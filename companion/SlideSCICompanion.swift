import AppKit
import Foundation

struct CompanionConfig {
  var helperCommand: String
  var helperScript: String?
  var helperPort: Int
  var webCommand: String
  var webScript: String?
  var webRoot: String?
  var webHost: String
  var webPort: Int
  var webCert: String?
  var webKey: String?
  var pollInterval: TimeInterval
  var shutdownGracePeriod: TimeInterval
  var powerpointBundleIdentifier: String
  var logPrefix: String
  var appleScriptFile: String?
}

enum ArgKey: String {
  case helperCommand = "--helper-command"
  case helperScript = "--helper-script"
  case helperPort = "--helper-port"
  case webCommand = "--web-command"
  case webScript = "--web-script"
  case webRoot = "--web-root"
  case webHost = "--host"
  case webPort = "--web-port"
  case webCert = "--web-cert"
  case webKey = "--web-key"
  case pollInterval = "--poll-interval"
  case shutdownGracePeriod = "--shutdown-grace-period"
  case powerpointBundleIdentifier = "--powerpoint-bundle-id"
  case logPrefix = "--log-prefix"
  case runAppleScriptFile = "--run-applescript-file"
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
    webCommand: values[ArgKey.webCommand.rawValue] ?? "/usr/bin/env",
    webScript: values[ArgKey.webScript.rawValue],
    webRoot: values[ArgKey.webRoot.rawValue],
    webHost: values[ArgKey.webHost.rawValue] ?? "127.0.0.1",
    webPort: Int(values[ArgKey.webPort.rawValue] ?? "18443") ?? 18443,
    webCert: values[ArgKey.webCert.rawValue],
    webKey: values[ArgKey.webKey.rawValue],
    pollInterval: TimeInterval(values[ArgKey.pollInterval.rawValue] ?? "3") ?? 3,
    shutdownGracePeriod: TimeInterval(values[ArgKey.shutdownGracePeriod.rawValue] ?? "5") ?? 5,
    powerpointBundleIdentifier: values[ArgKey.powerpointBundleIdentifier.rawValue] ?? "com.microsoft.Powerpoint",
    logPrefix: values[ArgKey.logPrefix.rawValue] ?? "[SlideSCI companion]",
    appleScriptFile: values[ArgKey.runAppleScriptFile.rawValue]
  )
}

final class ManagedProcess {
  let name: String
  private let executableURL: URL
  private let arguments: [String]
  private let logPrefix: String
  private let environment: [String: String]
  private let healthProbe: (() -> Bool)?
  private var process: Process?

  init(
    name: String,
    executableURL: URL,
    arguments: [String],
    logPrefix: String,
    environment: [String: String] = [:],
    healthProbe: (() -> Bool)? = nil
  ) {
    self.name = name
    self.executableURL = executableURL
    self.arguments = arguments
    self.logPrefix = logPrefix
    self.environment = environment
    self.healthProbe = healthProbe
  }

  func isRunning() -> Bool {
    guard let process else {
      return false
    }
    return process.isRunning
  }

  func ensureRunning() {
    if isRunning() {
      return
    }

    if let healthProbe, healthProbe() {
      log("检测到现有 \(name) 已可访问，跳过重复启动。")
      process = nil
      return
    }

    let nextProcess = Process()
    nextProcess.executableURL = executableURL
    nextProcess.arguments = arguments
    if !environment.isEmpty {
      nextProcess.environment = ProcessInfo.processInfo.environment.merging(environment) { _, new in new }
    }
    nextProcess.standardOutput = FileHandle.standardOutput
    nextProcess.standardError = FileHandle.standardError
    nextProcess.terminationHandler = { [weak self] terminatedProcess in
      self?.log("\(self?.name ?? "service") 已退出，状态码 \(terminatedProcess.terminationStatus)。")
      self?.process = nil
    }

    do {
      try nextProcess.run()
      process = nextProcess
      log("已启动 \(name) 进程。")
    } catch {
      process = nil
      log("启动 \(name) 失败：\(error.localizedDescription)")
    }
  }

  func stop(force: Bool = false) {
    guard let process, process.isRunning else {
      self.process = nil
      return
    }

    log("正在停止 \(name) 进程。")
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

    self.process = nil
  }

  private func log(_ message: String) {
    print("\(logPrefix) \(message)")
  }
}

final class CompanionSupervisor {
  private let config: CompanionConfig
  private let helperService: ManagedProcess
  private let webService: ManagedProcess
  private var helperStopDeadline: Date?

  init(config: CompanionConfig) {
    self.config = config

    let helperExecutable = URL(fileURLWithPath: config.helperCommand)
    self.helperService = ManagedProcess(
      name: "helper",
      executableURL: helperExecutable,
      arguments: Self.buildProcessArguments(command: config.helperCommand, script: config.helperScript, extraArguments: []),
      logPrefix: config.logPrefix,
      environment: ["SLIDESCI_APPLESCRIPT_RUNNER": CommandLine.arguments[0]],
      healthProbe: { Self.probeHealth(url: URL(string: "http://127.0.0.1:\(config.helperPort)/health")) }
    )

    let webExecutable = URL(fileURLWithPath: config.webCommand)
    self.webService = ManagedProcess(
      name: "本地任务窗格服务",
      executableURL: webExecutable,
      arguments: Self.buildProcessArguments(
        command: config.webCommand,
        script: config.webScript,
        extraArguments: [
          "--root", config.webRoot ?? "",
          "--host", config.webHost,
          "--port", String(config.webPort),
          "--cert", config.webCert ?? "",
          "--key", config.webKey ?? "",
          "--helper-host", "127.0.0.1",
          "--helper-port", String(config.helperPort),
        ]
      ),
      logPrefix: config.logPrefix,
      healthProbe: {
        Self.probeHealthWithCurl(urlString: "https://\(config.webHost):\(config.webPort)/health", allowInsecureTLS: true)
      }
    )
  }

  func tick() {
    ensureWebServiceRunning()
    if isPowerPointRunning() {
      helperStopDeadline = nil
      ensureHelperRunning()
      return
    }

    if !helperService.isRunning() {
      return
    }

    if helperStopDeadline == nil {
      helperStopDeadline = Date().addingTimeInterval(config.shutdownGracePeriod)
      log("PowerPoint 已退出，等待 \(Int(config.shutdownGracePeriod)) 秒后停止 helper。")
      return
    }

    if let helperStopDeadline, Date() >= helperStopDeadline {
      helperService.stop(force: true)
      self.helperStopDeadline = nil
    }
  }

  func shutdown() {
    helperService.stop(force: true)
    webService.stop(force: true)
  }

  private func ensureWebServiceRunning() {
    guard let webRoot = config.webRoot, !webRoot.isEmpty,
          let webScript = config.webScript, !webScript.isEmpty,
          let webCert = config.webCert, !webCert.isEmpty,
          let webKey = config.webKey, !webKey.isEmpty else {
      log("本地任务窗格服务配置不完整，已跳过启动。")
      return
    }

    guard FileManager.default.fileExists(atPath: webRoot),
          FileManager.default.fileExists(atPath: webScript),
          FileManager.default.fileExists(atPath: webCert),
          FileManager.default.fileExists(atPath: webKey) else {
      log("本地任务窗格服务所需文件缺失，已跳过启动。")
      return
    }

    webService.ensureRunning()
  }

  private func ensureHelperRunning() {
    guard let helperScript = config.helperScript, !helperScript.isEmpty else {
      log("helper 配置不完整，已跳过启动。")
      return
    }

    guard FileManager.default.fileExists(atPath: helperScript) else {
      log("helper 脚本不存在，已跳过启动。")
      return
    }

    helperService.ensureRunning()
  }

  private func isPowerPointRunning() -> Bool {
    let apps = NSRunningApplication.runningApplications(withBundleIdentifier: config.powerpointBundleIdentifier)
    return apps.contains(where: { !$0.isTerminated })
  }

  private func log(_ message: String) {
    print("\(config.logPrefix) \(message)")
  }

  private static func buildProcessArguments(command: String, script: String?, extraArguments: [String]) -> [String] {
    var arguments: [String] = []
    if let script, !script.isEmpty {
      if command == "/usr/bin/env" {
        arguments.append("node")
      }
      arguments.append(script)
    }
    return arguments + extraArguments.filter { !$0.isEmpty }
  }

  private static func probeHealth(url: URL?, session: URLSession = .shared) -> Bool {
    guard let url else {
      return false
    }

    let semaphore = DispatchSemaphore(value: 0)
    var success = false
    let task = session.dataTask(with: url) { _, response, _ in
      if let http = response as? HTTPURLResponse, (200..<500).contains(http.statusCode) {
        success = true
      }
      semaphore.signal()
    }
    task.resume()
    _ = semaphore.wait(timeout: .now() + 1.0)
    task.cancel()
    return success
  }

  private static func probeHealthWithCurl(urlString: String, allowInsecureTLS: Bool) -> Bool {
    let process = Process()
    process.executableURL = URL(fileURLWithPath: "/usr/bin/curl")
    var arguments = ["--silent", "--show-error", "--fail", "--max-time", "1"]
    if allowInsecureTLS {
      arguments.append("--insecure")
    }
    arguments.append(urlString)
    process.arguments = arguments
    process.standardOutput = Pipe()
    process.standardError = Pipe()

    do {
      try process.run()
      process.waitUntilExit()
      return process.terminationStatus == 0
    } catch {
      return false
    }
  }
}

func runAppleScriptFile(path: String) -> Int32 {
  do {
    let source = try String(contentsOfFile: path, encoding: .utf8)
    guard let script = NSAppleScript(source: source) else {
      FileHandle.standardError.write(Data("无法创建 NSAppleScript 实例。\n".utf8))
      return 1
    }

    var errorInfo: NSDictionary?
    let result = script.executeAndReturnError(&errorInfo)
    if let errorInfo {
      let message = errorInfo[NSAppleScript.errorMessage] as? String ?? errorInfo.description
      let number = errorInfo[NSAppleScript.errorNumber] as? Int ?? 1
      FileHandle.standardError.write(Data("\(message) (\(number))\n".utf8))
      return 1
    }

    if let stringValue = result.stringValue, !stringValue.isEmpty {
      print(stringValue)
    } else if result.descriptorType != typeNull {
      print(result.description)
    }
    return 0
  } catch {
    FileHandle.standardError.write(Data("\(error.localizedDescription)\n".utf8))
    return 1
  }
}

let config = parseArguments()
if let appleScriptFile = config.appleScriptFile, !appleScriptFile.isEmpty {
  exit(runAppleScriptFile(path: appleScriptFile))
}
let supervisor = CompanionSupervisor(config: config)

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
