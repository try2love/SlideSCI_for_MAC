import AppKit
import ApplicationServices
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
  var runnerQueueDir: String
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
  case requestAccessibilityPermission = "--request-accessibility-permission"
  case equationShortcut = "--equation-shortcut"
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
    appleScriptFile: values[ArgKey.runAppleScriptFile.rawValue],
    runnerQueueDir: "\(NSHomeDirectory())/Library/Application Support/SlideSCI/runner"
  )
}

func argumentValue(for key: ArgKey) -> String? {
  guard let index = CommandLine.arguments.firstIndex(of: key.rawValue),
        CommandLine.arguments.indices.contains(index + 1) else {
    return nil
  }
  return CommandLine.arguments[index + 1]
}

func hasFlag(_ key: ArgKey) -> Bool {
  CommandLine.arguments.contains(key.rawValue)
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

struct AppleScriptRunnerRequest: Codable {
  let id: String
  let scriptPath: String
}

struct AppleScriptRunnerResponse: Codable {
  let ok: Bool
  let stdout: String
  let stderr: String
}

func executeAppleScriptFile(path: String) -> (ok: Bool, output: String) {
  do {
    let source = try String(contentsOfFile: path, encoding: .utf8)
    guard let script = NSAppleScript(source: source) else {
      return (false, "无法创建 NSAppleScript 实例。")
    }

    var errorInfo: NSDictionary?
    let result = script.executeAndReturnError(&errorInfo)
    if let errorInfo {
      let message = errorInfo[NSAppleScript.errorMessage] as? String ?? errorInfo.description
      let number = errorInfo[NSAppleScript.errorNumber] as? Int ?? 1
      return (false, "\(message) (\(number))")
    }

    if let stringValue = result.stringValue, !stringValue.isEmpty {
      return (true, stringValue)
    }
    if result.descriptorType != typeNull {
      return (true, result.description)
    }
    return (true, "")
  } catch {
    return (false, error.localizedDescription)
  }
}

final class AppleScriptRunnerQueue {
  private let queueURL: URL
  private let encoder = JSONEncoder()
  private let decoder = JSONDecoder()
  private var processingIds = Set<String>()

  init(queueDir: String) {
    self.queueURL = URL(fileURLWithPath: queueDir)
    try? FileManager.default.createDirectory(at: queueURL, withIntermediateDirectories: true)
  }

  func processPendingRequests() {
    try? FileManager.default.createDirectory(at: queueURL, withIntermediateDirectories: true)
    guard let files = try? FileManager.default.contentsOfDirectory(
      at: queueURL,
      includingPropertiesForKeys: nil,
      options: [.skipsHiddenFiles]
    ) else {
      return
    }

    for requestURL in files where requestURL.lastPathComponent.hasSuffix(".request.json") {
      guard let request = readRequest(at: requestURL), !processingIds.contains(request.id) else {
        continue
      }
      processingIds.insert(request.id)
      handle(request: request, requestURL: requestURL)
      processingIds.remove(request.id)
    }
  }

  private func readRequest(at url: URL) -> AppleScriptRunnerRequest? {
    guard let data = try? Data(contentsOf: url) else {
      return nil
    }
    return try? decoder.decode(AppleScriptRunnerRequest.self, from: data)
  }

  private func handle(request: AppleScriptRunnerRequest, requestURL: URL) {
    let result = executeAppleScriptFile(path: request.scriptPath)
    let response = AppleScriptRunnerResponse(
      ok: result.ok,
      stdout: result.ok ? result.output : "",
      stderr: result.ok ? "" : result.output
    )
    let responseURL = queueURL.appendingPathComponent("\(request.id).response.json")
    if let data = try? encoder.encode(response) {
      try? data.write(to: responseURL, options: [.atomic])
    }
    try? FileManager.default.removeItem(at: requestURL)
  }
}

final class CompanionSupervisor {
  private let config: CompanionConfig
  private let helperService: ManagedProcess
  private let webService: ManagedProcess
  private let appleScriptRunnerQueue: AppleScriptRunnerQueue
  private var helperStopDeadline: Date?

  init(config: CompanionConfig) {
    self.config = config
    let accessibilityOptions = [kAXTrustedCheckOptionPrompt.takeUnretainedValue() as String: true] as CFDictionary
    if !AXIsProcessTrustedWithOptions(accessibilityOptions) {
      print("\(config.logPrefix) 已请求 macOS 辅助功能权限；授权后请完全退出并重新打开 PowerPoint。")
    }

    let helperExecutable = URL(fileURLWithPath: config.helperCommand)
    self.appleScriptRunnerQueue = AppleScriptRunnerQueue(queueDir: config.runnerQueueDir)
    self.helperService = ManagedProcess(
      name: "helper",
      executableURL: helperExecutable,
      arguments: Self.buildProcessArguments(command: config.helperCommand, script: config.helperScript, extraArguments: []),
      logPrefix: config.logPrefix,
      environment: [
        "SLIDESCI_APPLESCRIPT_RUNNER": CommandLine.arguments[0],
        "SLIDESCI_APPLESCRIPT_RUNNER_QUEUE": config.runnerQueueDir,
      ],
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
    appleScriptRunnerQueue.processPendingRequests()
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

  func processRunnerRequests() {
    appleScriptRunnerQueue.processPendingRequests()
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
  let result = executeAppleScriptFile(path: path)
  if result.ok {
    if !result.output.isEmpty {
      print(result.output)
    }
    return 0
  } else {
    FileHandle.standardError.write(Data("\(result.output)\n".utf8))
    return 1
  }
}

enum NativeEquationAutomationError: LocalizedError {
  case accessibilityPermissionMissing
  case powerpointNotRunning
  case focusedElementUnavailable
  case selectedTextRangeUnavailable
  case invalidPayload(String)
  case clipboardWriteFailed
  case eventDispatchFailed
  case accessibilityApiFailed(String)

  var errorDescription: String? {
    switch self {
    case .accessibilityPermissionMissing:
      return "macOS 未授予 SlideSCICompanion 辅助功能权限，无法驱动 PowerPoint 界面。"
    case .powerpointNotRunning:
      return "未检测到 Microsoft PowerPoint。"
    case .focusedElementUnavailable:
      return "无法进入文本编辑状态。请先选中文本框，再重试。"
    case .selectedTextRangeUnavailable:
      return "当前焦点元素不支持文本范围选择。请先进入文本框编辑状态，再重试。"
    case .invalidPayload(let message):
      return message
    case .clipboardWriteFailed:
      return "无法写入剪贴板内容。"
    case .eventDispatchFailed:
      return "无法发送键盘事件。"
    case .accessibilityApiFailed(let message):
      return message
    }
  }
}

func runningPowerPoint(bundleIdentifier: String) -> NSRunningApplication? {
  NSRunningApplication
    .runningApplications(withBundleIdentifier: bundleIdentifier)
    .first(where: { !$0.isTerminated })
}

func createKeyboardEvent(keyCode: CGKeyCode, keyDown: Bool, flags: CGEventFlags) throws -> CGEvent {
  guard let source = CGEventSource(stateID: .hidSystemState),
        let event = CGEvent(keyboardEventSource: source, virtualKey: keyCode, keyDown: keyDown) else {
    throw NativeEquationAutomationError.eventDispatchFailed
  }
  event.flags = flags
  return event
}

func sendKeyPress(keyCode: CGKeyCode, flags: CGEventFlags = []) throws {
  let keyDownEvent = try createKeyboardEvent(keyCode: keyCode, keyDown: true, flags: flags)
  let keyUpEvent = try createKeyboardEvent(keyCode: keyCode, keyDown: false, flags: flags)
  keyDownEvent.post(tap: .cghidEventTap)
  keyUpEvent.post(tap: .cghidEventTap)
}

func copyAttributeValue(_ element: AXUIElement, attribute: String) throws -> CFTypeRef {
  var value: CFTypeRef?
  let error = AXUIElementCopyAttributeValue(element, attribute as CFString, &value)
  guard error == .success, let value else {
    throw NativeEquationAutomationError.accessibilityApiFailed("无法读取辅助功能属性 \(attribute)：\(error.rawValue)")
  }
  return value
}

func currentFocusedElement() -> AXUIElement? {
  let systemWide = AXUIElementCreateSystemWide()
  var value: CFTypeRef?
  guard AXUIElementCopyAttributeValue(systemWide, kAXFocusedUIElementAttribute as CFString, &value) == .success,
        let value else {
    return nil
  }
  return (value as! AXUIElement)
}

func focusedEditableElement(bundleIdentifier: String) throws -> AXUIElement {
  guard AXIsProcessTrusted() else {
    throw NativeEquationAutomationError.accessibilityPermissionMissing
  }
  guard let app = runningPowerPoint(bundleIdentifier: bundleIdentifier) else {
    throw NativeEquationAutomationError.powerpointNotRunning
  }

  app.activate(options: [.activateIgnoringOtherApps])
  usleep(150_000)

  if let focused = currentFocusedElement() {
    var rangeValue: CFTypeRef?
    if AXUIElementCopyAttributeValue(focused, kAXSelectedTextRangeAttribute as CFString, &rangeValue) == .success {
      return focused
    }
  }

  try sendKeyPress(keyCode: 36)
  usleep(150_000)

  guard let focused = currentFocusedElement() else {
    throw NativeEquationAutomationError.focusedElementUnavailable
  }
  var rangeValue: CFTypeRef?
  guard AXUIElementCopyAttributeValue(focused, kAXSelectedTextRangeAttribute as CFString, &rangeValue) == .success else {
    throw NativeEquationAutomationError.selectedTextRangeUnavailable
  }
  return focused
}

func setSelectedTextRange(_ element: AXUIElement, start: Int, length: Int) throws {
  var range = CFRange(location: start, length: length)
  guard let rangeValue = AXValueCreate(.cfRange, &range) else {
    throw NativeEquationAutomationError.accessibilityApiFailed("无法创建文本范围描述。")
  }
  let error = AXUIElementSetAttributeValue(element, kAXSelectedTextRangeAttribute as CFString, rangeValue)
  guard error == .success else {
    throw NativeEquationAutomationError.accessibilityApiFailed("无法设置文本范围：\(error.rawValue)")
  }
}

func requestAccessibilityPermission() -> Int32 {
  let options = [kAXTrustedCheckOptionPrompt.takeUnretainedValue() as String: true] as CFDictionary
  let trusted = AXIsProcessTrustedWithOptions(options)
  print(trusted ? "trusted" : "prompted")
  return 0
}

func runEquationShortcut(bundleIdentifier: String) -> Int32 {
  do {
    guard let app = runningPowerPoint(bundleIdentifier: bundleIdentifier) else {
      throw NativeEquationAutomationError.powerpointNotRunning
    }
    app.activate(options: [.activateIgnoringOtherApps])
    usleep(120_000)
    try sendKeyPress(keyCode: 24, flags: .maskAlternate)
    usleep(180_000)
    print("equation-insert")
    return 0
  } catch {
    FileHandle.standardError.write(Data("\((error as? LocalizedError)?.errorDescription ?? error.localizedDescription)\n".utf8))
    return 1
  }
}

let config = parseArguments()
if let appleScriptFile = config.appleScriptFile, !appleScriptFile.isEmpty {
  exit(runAppleScriptFile(path: appleScriptFile))
}
if hasFlag(.requestAccessibilityPermission) {
  exit(requestAccessibilityPermission())
}
if hasFlag(.equationShortcut) {
  exit(runEquationShortcut(bundleIdentifier: config.powerpointBundleIdentifier))
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

let runnerTimer = DispatchSource.makeTimerSource()
runnerTimer.schedule(deadline: .now(), repeating: 0.05)
runnerTimer.setEventHandler {
  supervisor.processRunnerRequests()
}
runnerTimer.resume()

RunLoop.main.run()
