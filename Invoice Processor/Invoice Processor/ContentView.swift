import SwiftUI
import UniformTypeIdentifiers

struct ContentView: View {
    @State private var droppedURL: URL?
    @State private var isProcessing = false
    @State private var isRegistering = false
    @State private var result: ProcessResult?
    @State private var errorMessage: String?
    @State private var isTargeted = false
    @State private var skuMasterName: String = loadSkuMasterName()
    @State private var registerSuccess: String?

    var body: some View {
        VStack(spacing: 0) {
            headerView
            Divider()
            contentView
                .padding(24)
                .frame(maxWidth: .infinity, maxHeight: .infinity)
                .background(Color(NSColor.controlBackgroundColor))
        }
    }

    // MARK: - Header
    private var headerView: some View {
        HStack {
            Image(systemName: "doc.viewfinder")
                .font(.title2)
                .foregroundStyle(.blue)
            Text("Invoice Processor")
                .font(.system(size: 15, weight: .semibold))
            Spacer()
            if droppedURL != nil || result != nil {
                Button(action: reset) {
                    Image(systemName: "arrow.counterclockwise")
                        .foregroundStyle(.secondary)
                }
                .buttonStyle(.plain)
                .help("Reset")
            }
        }
        .padding(.horizontal, 20)
        .padding(.vertical, 12)
        .background(.ultraThinMaterial)
    }

    // MARK: - Content
    @ViewBuilder
    private var contentView: some View {
        if let result = result {
            resultView(result)
        } else if isProcessing {
            processingView
        } else if isRegistering {
            registeringView
        } else {
            dropZoneView
        }
    }

    // MARK: - Drop Zone
    private var dropZoneView: some View {
        VStack(spacing: 16) {
            ZStack {
                RoundedRectangle(cornerRadius: 16)
                    .strokeBorder(
                        isTargeted ? Color.blue : Color.gray.opacity(0.25),
                        style: StrokeStyle(lineWidth: isTargeted ? 2.5 : 2, dash: [8, 6])
                    )
                    .background(
                        RoundedRectangle(cornerRadius: 16)
                            .fill(isTargeted ? Color.blue.opacity(0.06) : Color.clear)
                    )
                    .frame(height: 240)

                VStack(spacing: 14) {
                    Image(systemName: "arrow.down.doc")
                        .font(.system(size: 36))
                        .foregroundStyle(isTargeted ? .blue : .secondary)

                    VStack(spacing: 4) {
                        Text("Drop invoice PDF here")
                            .font(.system(size: 15, weight: .medium))
                        Text("or drop SKU master Excel to register")
                            .font(.system(size: 12))
                            .foregroundStyle(.secondary)
                    }

                    HStack(spacing: 6) {
                        Image(systemName: "doc.text.magnifyingglass")
                            .font(.caption)
                            .foregroundStyle(.tertiary)
                        Text("or click to select a file")
                            .font(.system(size: 12))
                            .foregroundStyle(.tertiary)
                    }
                }
                .foregroundStyle(.primary)
            }
            .onDrop(of: [.fileURL], isTargeted: $isTargeted) { providers in
                _ = handleDrop(providers)
                return true
            }
            .onTapGesture {
                selectFile()
            }

            // SKU master info
            skuMasterInfoView
        }
    }

    // MARK: - SKU Master Info
    private var skuMasterInfoView: some View {
        HStack(spacing: 10) {
            Image(systemName: "tablecells.badge.ellipsis")
                .font(.caption)
                .foregroundStyle(.blue)

            if skuMasterName.isEmpty {
                Text("No SKU master registered")
                    .font(.system(size: 12))
                    .foregroundStyle(.red)
            } else {
                Text("SKU master: \(skuMasterName)")
                    .font(.system(size: 12, weight: .medium))
                    .foregroundStyle(.secondary)
                    .lineLimit(1)
                    .truncationMode(.middle)
            }

            Spacer()

            Button("Change...") {
                selectSkuMaster()
            }
            .buttonStyle(.plain)
            .font(.system(size: 11, weight: .medium))
            .foregroundStyle(.blue)
            .help("Register a different SKU master file")
        }
        .padding(.horizontal, 14)
        .padding(.vertical, 10)
        .background(
            RoundedRectangle(cornerRadius: 8)
                .fill(Color(NSColor.controlBackgroundColor))
                .overlay(
                    RoundedRectangle(cornerRadius: 8)
                        .stroke(Color.gray.opacity(0.15), lineWidth: 1)
                )
        )
    }

    // MARK: - Registering
    private var registeringView: some View {
        VStack(spacing: 24) {
            Image(systemName: "checkmark.circle.fill")
                .font(.system(size: 40))
                .foregroundStyle(.green)

            if let success = registerSuccess {
                Text(success)
                    .font(.system(size: 14, weight: .medium))
                    .multilineTextAlignment(.center)
            }

            Text("SKU master is ready to use")
                .font(.system(size: 12))
                .foregroundStyle(.secondary)

            HStack(spacing: 12) {
                Button(action: reset) {
                    Label("Process a PDF", systemImage: "doc.viewfinder")
                        .frame(minWidth: 130)
                }
                .buttonStyle(.borderedProminent)
                .tint(.blue)

                Button(action: selectSkuMaster) {
                    Label("Change", systemImage: "arrow.triangle.swap")
                        .frame(minWidth: 100)
                }
                .buttonStyle(.bordered)
            }
            .controlSize(.regular)
        }
    }

    // MARK: - Processing
    private var processingView: some View {
        VStack(spacing: 24) {
            ProgressView()
                .scaleEffect(1.2)
                .controlSize(.large)

            if let url = droppedURL {
                HStack(spacing: 8) {
                    Image(systemName: "doc.text")
                        .foregroundStyle(.secondary)
                    Text(url.lastPathComponent)
                        .font(.system(size: 13, weight: .medium))
                        .lineLimit(1)
                        .truncationMode(.middle)
                }
                .foregroundStyle(.secondary)
            }

            HStack(spacing: 4) {
                Text("Using:")
                    .foregroundStyle(.tertiary)
                Text(skuMasterName)
                    .fontWeight(.medium)
                    .foregroundStyle(.secondary)
            }
            .font(.system(size: 12))
        }
    }

    // MARK: - Result
    private func resultView(_ result: ProcessResult) -> some View {
        let isBatch = result.fileResults.count > 1
        return VStack(spacing: 12) {
            if isBatch {
                batchTabView(result)
            } else {
                singleResultView(result, fileResult: nil)
            }

            HStack(spacing: 12) {
                Button(action: openFolder) {
                    Label("Show in Finder", systemImage: "folder")
                        .frame(minWidth: 130)
                }
                .buttonStyle(.borderedProminent)
                .tint(.blue)

                Button(action: processAnother) {
                    Label("Process Another", systemImage: "plus")
                        .frame(minWidth: 130)
                }
                .buttonStyle(.bordered)
            }
            .controlSize(.regular)
        }
    }

    private func batchTabView(_ result: ProcessResult) -> some View {
        VStack(spacing: 8) {
            TabView {
                ForEach(result.fileResults) { file in
                    singleResultView(result, fileResult: file)
                        .tabItem {
                            Label(shortName(file.filename), systemImage: "doc")
                        }
                }
            }
        }
    }

    private func shortName(_ name: String) -> String {
        // "SG106768248 CI.pdf" → "SG106768248"
        let trimmed = name.hasSuffix(".pdf") ? String(name.dropLast(4)) : name
        if trimmed.hasSuffix(" CI") { return String(trimmed.dropLast(3)) }
        return trimmed
    }

    private func singleResultView(_ result: ProcessResult, fileResult: FileResult?) -> some View {
        let isFile = fileResult != nil
        let fr = fileResult

        return VStack(spacing: 16) {
            VStack(spacing: 8) {
                // File name header (only in per-file tabs)
                if isFile, let name = fr?.filename {
                    HStack(spacing: 6) {
                        Image(systemName: "doc.text")
                            .foregroundStyle(.blue)
                            .font(.caption)
                        Text(name)
                            .font(.system(size: 12, weight: .medium))
                            .lineLimit(1)
                            .truncationMode(.middle)
                    }
                    .foregroundStyle(.secondary)
                }

                VStack(spacing: 8) {
                    if isFile {
                        summaryRow(icon: "building.storefront", label: "Retailer", value: fr?.retailerLabel ?? result.retailerLabel)
                        Divider()
                        summaryRow(icon: "number", label: "Items", value: "\(fr?.itemCount ?? result.itemCount)")
                        summaryRow(icon: "link.badge.plus", label: "SKU Matched", value: "\(fr?.matchedCount ?? result.matchedCount)/\(fr?.itemCount ?? result.itemCount)")
                        summaryRow(icon: "shippingbox", label: "Quantity", value: formattedNumber(fr?.totalQty ?? result.totalQty))
                        summaryRow(icon: "dollarsign", label: "Total Price USD", value: formattedCurrency(fr?.totalAmount ?? result.totalAmount))
                    } else {
                        summaryRow(icon: "building.storefront", label: "Retailer", value: result.retailerLabel)
                        if result.fileResults.count > 1 {
                            summaryRow(icon: "doc.on.doc", label: "Files", value: "\(result.fileResults.count)")
                        }
                        Divider()
                        summaryRow(icon: "number", label: "Total Items", value: "\(result.itemCount)")
                        summaryRow(icon: "link.badge.plus", label: "SKU Matched", value: "\(result.matchedCount)/\(result.itemCount)")
                        summaryRow(icon: "shippingbox", label: "Total Quantity", value: formattedNumber(result.totalQty))
                        summaryRow(icon: "dollarsign", label: "Total Price USD", value: formattedCurrency(result.totalAmount))
                    }
                }
                .padding(14)
                .background(
                    RoundedRectangle(cornerRadius: 10)
                        .fill(Color(NSColor.controlBackgroundColor))
                        .shadow(color: .black.opacity(0.05), radius: 4, y: 2)
                )

                // Open button (per-file)
                if isFile, let path = fr?.outputPath {
                    Button(action: { NSWorkspace.shared.open(URL(fileURLWithPath: path)) }) {
                        Label("Open Excel", systemImage: "tablecells")
                            .frame(minWidth: 140)
                    }
                    .buttonStyle(.borderedProminent)
                    .tint(.blue)
                    .controlSize(.regular)
                }
            }
            .padding(.horizontal, 16)
        }
    }

    private func summaryRow(icon: String, label: String, value: String) -> some View {
        HStack {
            Image(systemName: icon)
                .font(.system(size: 12))
                .foregroundStyle(.blue)
                .frame(width: 18)
            Text(label)
                .font(.system(size: 13))
                .foregroundStyle(.secondary)
            Spacer()
            Text(value)
                .font(.system(size: 13, weight: .semibold, design: .monospaced))
                .foregroundStyle(.primary)
        }
    }

    // MARK: - Actions
    private func processAnother() {
        reset()
        selectFile()
    }

    private func reset() {
        droppedURL = nil
        result = nil
        errorMessage = nil
        isProcessing = false
        isRegistering = false
        registerSuccess = nil
    }

    private func selectFile() {
        let panel = NSOpenPanel()
        panel.allowedContentTypes = [.pdf]
        panel.allowsMultipleSelection = true
        panel.canChooseDirectories = false
        panel.message = "Select invoice PDF(s)"
        panel.prompt = "Process"

        guard panel.runModal() == .OK else { return }
        let pdfs = panel.urls.filter { $0.pathExtension.lowercased() == "pdf" }
        if pdfs.count == 1 {
            processPDF(pdfs[0])
        } else if pdfs.count > 1 {
            processBatch(pdfs)
        }
    }

    private func selectSkuMaster() {
        let panel = NSOpenPanel()
        panel.allowedContentTypes = [UTType(filenameExtension: "xlsx")].compactMap { $0 }
        panel.allowsMultipleSelection = false
        panel.canChooseDirectories = false
        panel.message = "Select a SKU master Excel file"
        panel.prompt = "Register"

        guard panel.runModal() == .OK, let url = panel.url else { return }
        registerSkuMaster(url)
    }

    private func handleDrop(_ providers: [NSItemProvider]) -> Bool {
        var urls: [URL] = []
        let group = DispatchGroup()

        for provider in providers {
            group.enter()
            provider.loadItem(forTypeIdentifier: UTType.fileURL.identifier, options: nil) { item, error in
                if let data = item as? Data,
                   let url = URL(dataRepresentation: data, relativeTo: nil) {
                    urls.append(url)
                }
                group.leave()
            }
        }

        group.notify(queue: .main) {
            let pdfs = urls.filter { $0.pathExtension.lowercased() == "pdf" }
            let xlsx = urls.filter { $0.pathExtension.lowercased() == "xlsx" }
            if let sku = xlsx.first {
                self.registerSkuMaster(sku)
            } else if pdfs.count == 1 {
                self.processPDF(pdfs[0])
            } else if pdfs.count > 1 {
                self.processBatch(pdfs)
            }
        }
        return true
    }

    private func registerSkuMaster(_ url: URL) {
        isRegistering = true
        isProcessing = false
        result = nil
        registerSuccess = nil

        DispatchQueue.global(qos: .userInitiated).async {
            let (name, error) = registerSkuMasterFile(fileURL: url)
            DispatchQueue.main.async {
                if let name = name {
                    self.skuMasterName = name
                    self.registerSuccess = "Registered: \(name)"
                } else if let error = error {
                    self.isRegistering = false
                    self.errorMessage = error
                }
            }
        }
    }

    private func processPDF(_ url: URL) {
        droppedURL = url
        isProcessing = true
        result = nil
        errorMessage = nil

        DispatchQueue.global(qos: .userInitiated).async {
            let (result, error) = runProcessor(pdfPath: url.path)
            DispatchQueue.main.async {
                self.isProcessing = false
                if let result = result {
                    self.result = result
                    // Refresh SKU master name in case it changed
                    self.skuMasterName = loadSkuMasterName()
                } else if let error = error {
                    self.errorMessage = error
                }
            }
        }
    }

    private func processBatch(_ urls: [URL]) {
        isProcessing = true
        result = nil
        errorMessage = nil

        DispatchQueue.global(qos: .userInitiated).async {
            var totalItems = 0, totalMatched = 0, totalQty = 0
            var totalPrice = 0.0
            var fileCount = 0
            var firstResult: ProcessResult?
            var lastOutputPath: String?
            var fileResults: [FileResult] = []
            var batchRetailer = ""

            for (_, url) in urls.enumerated() {
                DispatchQueue.main.async {
                    self.droppedURL = url
                    self.isProcessing = true
                }
                let (res, _) = runProcessor(pdfPath: url.path)
                if let res = res {
                    totalItems += res.itemCount
                    totalMatched += res.matchedCount
                    totalQty += res.totalQty
                    totalPrice += res.totalAmount
                    fileCount += 1
                    if firstResult == nil { firstResult = res }
                    if batchRetailer.isEmpty { batchRetailer = res.retailerLabel }
                    lastOutputPath = res.outputPath
                    DispatchQueue.main.async {
                        self.skuMasterName = loadSkuMasterName()
                    }
                    if let path = res.outputPath {
                        fileResults.append(FileResult(
                            filename: res.filename ?? url.lastPathComponent,
                            itemCount: res.itemCount,
                            matchedCount: res.matchedCount,
                            totalQty: res.totalQty,
                            totalAmount: res.totalAmount,
                            outputPath: path,
                            retailerLabel: res.retailerLabel
                        ))
                    }
                }
            }

            let combinedResult = ProcessResult(
                itemCount: totalItems,
                matchedCount: totalMatched,
                totalQty: totalQty,
                totalAmount: totalPrice,
                retailerLabel: batchRetailer,
                filename: fileCount > 1 ? "\(fileCount) PDFs" : firstResult?.filename,
                outputPath: lastOutputPath,
                fileResults: fileResults
            )

            DispatchQueue.main.async {
                self.isProcessing = false
                self.result = combinedResult
            }
        }
    }

    private func openExcel() {
        guard let r = result, let path = r.outputPath else { return }
        NSWorkspace.shared.open(URL(fileURLWithPath: path))
    }

    private func openFolder() {
        guard let r = result, let path = r.outputPath else { return }
        NSWorkspace.shared.activateFileViewerSelecting([URL(fileURLWithPath: path)])
    }
}

// MARK: - SKU Master Registration
func registerSkuMasterFile(fileURL: URL) -> (String?, String?) {
    let destDir = "\(NSHomeDirectory())/hermes work/po-process/sku-master"
    let configPath = "\(NSHomeDirectory())/hermes work/po-process/config.json"

    // Create directory
    try? FileManager.default.createDirectory(atPath: destDir, withIntermediateDirectories: true)

    // Preserve original filename with timestamp
    let srcName = fileURL.lastPathComponent
    let nameOnly = (srcName as NSString).deletingPathExtension
    let now = DateFormatter()
    now.dateFormat = "yyyyMM"
    let timestamp = now.string(from: Date())
    let destName: String
    if nameOnly.hasPrefix("SKU master file") {
        // Keep original name if it already has the pattern
        destName = srcName
    } else {
        destName = "SKU master file_\(timestamp).xlsx"
    }
    let destPath = "\(destDir)/\(destName)"

    // Remove old SKU master files
    if let files = try? FileManager.default.contentsOfDirectory(atPath: destDir) {
        for f in files where f.hasPrefix("SKU master file") && f.hasSuffix(".xlsx") {
            try? FileManager.default.removeItem(atPath: "\(destDir)/\(f)")
        }
    }

    // Copy new file
    do {
        if FileManager.default.fileExists(atPath: destPath) {
            try FileManager.default.removeItem(atPath: destPath)
        }
        try FileManager.default.copyItem(atPath: fileURL.path, toPath: destPath)
    } catch {
        return (nil, "Failed to copy file: \(error.localizedDescription)")
    }

    // Update config
    let config: [String: Any] = ["sku_master_path": destPath]
    if let data = try? JSONSerialization.data(withJSONObject: config, options: [.prettyPrinted]) {
        try? data.write(to: URL(fileURLWithPath: configPath))
    }

    return (destName, nil)
}

func loadSkuMasterName() -> String {
    let configPath = "\(NSHomeDirectory())/hermes work/po-process/config.json"
    guard let data = try? Data(contentsOf: URL(fileURLWithPath: configPath)),
          let config = try? JSONSerialization.jsonObject(with: data) as? [String: String],
          let path = config["sku_master_path"],
          FileManager.default.fileExists(atPath: path) else {
        // Fallback: check Downloads
        let downloads = "\(NSHomeDirectory())/Downloads"
        if let files = try? FileManager.default.contentsOfDirectory(atPath: downloads) {
            let candidates = files.filter { $0.hasPrefix("SKU master file") && $0.hasSuffix(".xlsx") }
            if let latest = candidates.sorted().last {
                return latest
            }
        }
        return ""
    }
    return (path as NSString).lastPathComponent
}

// MARK: - Processor
struct FileResult: Identifiable {
    let id = UUID()
    let filename: String
    let itemCount: Int
    let matchedCount: Int
    let totalQty: Int
    let totalAmount: Double
    let outputPath: String
    let retailerLabel: String
}

struct ProcessResult {
    let itemCount: Int
    let matchedCount: Int
    let totalQty: Int
    let totalAmount: Double
    let retailerLabel: String
    let filename: String?
    let outputPath: String?
    let fileResults: [FileResult]  // individual file results for batch mode
}

func runProcessor(pdfPath: String) -> (ProcessResult?, String?) {
    let scriptPath = Bundle.main.path(forResource: "process_invoice", ofType: "py")
                     ?? "\(NSHomeDirectory())/hermes work/po-process/process_invoice.py"

    let task = Process()
    task.executableURL = URL(fileURLWithPath: "/usr/bin/env")
    task.arguments = ["python3", scriptPath, pdfPath]

    let outputPipe = Pipe()
    let errorPipe = Pipe()
    task.standardOutput = outputPipe
    task.standardError = errorPipe

    do {
        try task.run()
        task.waitUntilExit()

        let outputData = outputPipe.fileHandleForReading.readDataToEndOfFile()
        let output = String(data: outputData, encoding: .utf8) ?? ""

        guard task.terminationStatus == 0 else {
            let err = String(data: errorPipe.fileHandleForReading.readDataToEndOfFile(), encoding: .utf8) ?? "Unknown error"
            return (nil, err)
        }

        var itemCount = 0, matchedCount = 0, totalQty = 0
        var totalAmount = 0.0, outputPath: String? = nil
        var retailerLabel = ""

        for line in output.components(separatedBy: .newlines) {
            let trimmed = line.trimmingCharacters(in: .whitespaces)
            if trimmed.contains("리테일러:") {
                if let range = trimmed.range(of: "리테일러: ") {
                    retailerLabel = String(trimmed[range.upperBound...]).trimmingCharacters(in: .whitespaces)
                }
            }
            if trimmed.contains("품목 처리") {
                itemCount = Int(trimmed.components(separatedBy: CharacterSet.decimalDigits.inverted).joined()) ?? 0
            }
            if trimmed.contains("매칭") {
                let nums = trimmed.components(separatedBy: CharacterSet.decimalDigits.inverted).filter { !$0.isEmpty }
                if nums.count >= 2 {
                    matchedCount = Int(nums[0]) ?? 0
                }
            }
            if trimmed.contains("총 수량") {
                let clean = trimmed.replacingOccurrences(of: ",", with: "")
                let nums = clean.components(separatedBy: CharacterSet.decimalDigits.inverted).filter { !$0.isEmpty }
                totalQty = Int(nums.last ?? "") ?? 0
            }
            if trimmed.contains("총 금액") {
                // Extract number: find "$\d+[\d,]*\.?\d*" pattern
                if let dollarRange = trimmed.range(of: "$") {
                    let afterDollar = trimmed[dollarRange.upperBound...]
                    let numStr = afterDollar.trimmingCharacters(in: .whitespaces)
                    let clean = numStr.replacingOccurrences(of: ",", with: "")
                    if let val = Double(clean) {
                        totalAmount = val
                    }
                }
            }
            if trimmed.contains("결과:") {
                if let range = trimmed.range(of: "/Users/") {
                    outputPath = String(trimmed[range.lowerBound...]).trimmingCharacters(in: .whitespaces)
                }
            }
        }

        let pdfURL = URL(fileURLWithPath: pdfPath)
        let filename = pdfURL.lastPathComponent

        if let path = outputPath {
            for _ in 0..<10 {
                if FileManager.default.fileExists(atPath: path) { break }
                Thread.sleep(forTimeInterval: 0.1)
            }
        }

        return (ProcessResult(
            itemCount: itemCount,
            matchedCount: matchedCount,
            totalQty: totalQty,
            totalAmount: totalAmount,
            retailerLabel: retailerLabel,
            filename: filename,
            outputPath: outputPath,
            fileResults: []
        ), nil)
    } catch {
        return (nil, error.localizedDescription)
    }
}

// MARK: - Helpers
func formattedNumber(_ n: Int) -> String {
    let formatter = NumberFormatter()
    formatter.numberStyle = .decimal
    formatter.groupingSeparator = ","
    return formatter.string(from: NSNumber(value: n)) ?? "\(n)"
}

func formattedCurrency(_ n: Double) -> String {
    let formatter = NumberFormatter()
    formatter.numberStyle = .currency
    formatter.currencyCode = "USD"
    formatter.maximumFractionDigits = 2
    return formatter.string(from: NSNumber(value: n)) ?? "$\(n)"
}
