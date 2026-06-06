import SwiftUI

@main
struct InvoiceProcessorApp: App {
    var body: some Scene {
        Window("Invoice Processor", id: "main") {
            ContentView()
                .frame(minWidth: 520, minHeight: 460)
        }
        .windowResizability(.contentSize)
        .windowStyle(.hiddenTitleBar)
    }
}
