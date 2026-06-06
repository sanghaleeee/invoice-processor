// swift-tools-version:5.9
import PackageDescription

let package = Package(
    name: "Invoice Processor",
    platforms: [.macOS(.v14)],
    targets: [
        .executableTarget(
            name: "Invoice Processor",
            resources: [
                .copy("process_invoice.py"),
                .copy("SKU master file_2026.04.xlsx")
            ]
        )
    ]
)
