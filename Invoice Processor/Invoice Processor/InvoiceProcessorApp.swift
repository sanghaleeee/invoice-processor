import SwiftUI
import UniformTypeIdentifiers

@main
struct InvoiceProcessorApp: App {
    @NSApplicationDelegateAdaptor(AppDelegate.self) var appDelegate

    var body: some Scene {
        Window("Invoice Processor", id: "main") {
            ContentView()
                .frame(minWidth: 520, minHeight: 460)
                .onOpenURL { url in
                    // Handle file:// URLs (drag onto Dock icon)
                    NotificationCenter.default.post(name: .handleDroppedFile, object: url)
                }
        }
        .windowResizability(.contentSize)
        .windowStyle(.hiddenTitleBar)
        .commands {
            CommandGroup(after: .newItem) {
                Button("Open PDF…") {
                    NotificationCenter.default.post(name: .openFileDialog, object: nil)
                }
                .keyboardShortcut("o")
            }
        }
    }
}

class AppDelegate: NSObject, NSApplicationDelegate {
    func application(_ application: NSApplication, open urls: [URL]) {
        // Handle files dropped onto Dock/App icon
        for url in urls {
            NotificationCenter.default.post(name: .handleDroppedFile, object: url)
        }
    }
}

extension Notification.Name {
    static let handleDroppedFile = Notification.Name("handleDroppedFile")
    static let openFileDialog = Notification.Name("openFileDialog")
}
