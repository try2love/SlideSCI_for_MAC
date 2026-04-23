import AppKit
import Foundation

struct CompanionConfig {
  var helperCommand: String
  var helperScript: String?
  var helperPort: Int
  var pollInterval: TimeInterval
  var shutdownGracePeriod: TimeInterval
  var powerpointBundleIdentifier: String
  var logPrefix: String
}

enum ArgKey: String {
  case helperCommand = "--helper-command"
  case helperScript = "--helper-script"
  case helperPort = "--helper-port"
  case pollInterval = "--poll-interval"
  case shutdownGracePeriod = "--shutdown-grace-period"
  case powerpointBundleIdentifier = "--powerpoint-bundle-id"
  case logPrefix = "--log-prefix"
}

func parseArguments() -> CompanionConfig {
  var values: [String: String] = [:]
  var index = 1
  while index < CommandLine.arguments.count {
    let key = CommandLine.arguments[index]
    if let argKey = ArgKey(rawValue: key), index + 1 < CommandLine.arguments.count {
      values[argKey.rawValue] = CommandLine.arguments[index + 1]
      index += 2
      continue
    }
    index += 1
  }

  return CompanionConfig(
    helperCommand: values[ArgKey.helperCommand.rawValue] ?? "/usr/bin/env",
    helperScript: values[ArgKey.helperScript.rawValue],
    helperPort: Int(values[ArgKey.helperPort.rawValue] ?? "17926") ?? 17926,
    pollInterval: TimeInterval(values[ArgKey.pollInterval.rawValue] ?? "3") ?? 3,
    shutdownGracePeriod: TimeInterval(values[ArgKey.shutdownGracePeriod.rawValue] ?? "5") ?? 5,
    powerpointBundleIdentifier: values[ArgKey.powerpointBundleIdentifier.rawValue] ?? "com.microsoft.Powerpoint",
    logPrefix: values[ArgKey.logPrefix.rawValue] ?? "[SlideSCI companion]"
  )
}

final class HelperSupervisor {
  private let config: CompanionConfig
  private var helperProcess: Process?
  private var stopDeadline: Date?

  init(config: CompanionConfig) {
    self.config = config
  }

  func tick() {
    let powerpointRunning = isPowerPointRunning()
    if powerpointRunning {
      stopDeadline = nil
      ensureHelperRunning()
      return
    }

    if helperProcess == nil {
      return
    }

    if stopDeadline == nil {
      stopDeadline = Date().addingTimeInterval(config.shutdownGracePeriod)
      log("PowerPoint 已退出，等待 \(Int(config.shutdownGracePeriod)) 秒后停止 helper。")
      return
    }

    if let deadline = stopDeadline, Date() >= deadline {
      stopHelper()
      stopDeadline = nil
    }
  }

  func shutdown() {
    stopHelper(force: true)
  }

  private func isPowerPointRunning() -> Bool {
    NSWorkspace.shared.runningApplications.contains { app in
      app.bundleIdentifier == config.powerpointBundleIdentifier && !app.isTerminated
    }
  }

  private func helperProcessIsAlive() -> Bool {
    guard let process = helperProcess else {
      return false
    }
    return process.isRunning
  }

  private func ensureHelperRunning() {
    if helperProcessIsAlive() {
      return
    }

    if isHelperPortResponding() {
      log("检测到现有 helper 已监听 127.0.0.1:\(config.helperPort)，跳过重复启动。")
      helperProcess = nil
      return
    }

    let process = Process()
    process.executableURL = URL(fileURLWithPath: config.helperCommand)

    if let helperScript = config.helperScript, !helperScript.isEmpty {
      if config.helperCommand == "/usr/bin/env" {
        process.arguments = ["node", helperScript]
      } else {
        process.arguments = [helperScript]
      }
    } else {
      process.arguments = []
    }

    let output = Pipe()
    process.standardOutput = output
    process.standardError = output
    process.terminationHandler = { [weak self] terminatedProcess in
      self?.log("helper 已退出，状态码 \(terminatedProcess.terminationStatus)。")
      self?.helperProcess = nil
    }

    do {
      try process.run()
      helperProcess = process
      log("已启动 helper 进程。")
    } catch {
      helperProcess = nil
      log("启动 helper 失败：\(error.localizedDescription)")
    }
  }

  private func stopHelper(force: Bool = false) {
    guard let process = helperProcess, process.isRunning else {
      helperProcess = nil
      return
    }

    log("正在停止 helper 进程。")
    process.terminate()

    if force {
      let deadline = Date().addingTimeInterval(1.5)
      while process.isRunning && Date() < deadline {
        RunLoop.current.run(mode: .default, before: Date().addingTimeInterval(0.05))
      }
      if process.isRunning {
        kill(process.processIdentifier, SIGKILL)
      }
    }

    helperProcess = nil
  }

  private func isHelperPortResponding() -> Bool {
    guard let url = URL(string: "http://127.0.0.1:\(config.helperPort)/health") else {
      return false
    }

    let semaphore = DispatchSemaphore(value: 0)
    var success = false
    let task = URLSession.shared.dataTask(with: url) { _, response, _ in
      if let http = response as? HTTPURLResponse, (200..<500).contains(http.statusCode) {
        success = true
      }
      semaphore.signal()
    }
    task.resume()
    _ = semaphore.wait(timeout: .now() + 0.8)
    task.cancel()
    return success
  }

  private func log(_ message: String) {
    print("\(config.logPrefix) \(message)")
  }
}

let config = parseArguments()
let supervisor = HelperSupervisor(config: config)

signal(SIGTERM) { _ in
  supervisor.shutdown()
  exit(0)
}

signal(SIGINT) { _ in
  supervisor.shutdown()
  exit(0)
}

let timer = DispatchSource.makeTimerSource()
timer.schedule(deadline: .now(), repeating: config.pollInterval)
timer.setEventHandler {
  supervisor.tick()
}
timer.resume()

RunLoop.main.run()
