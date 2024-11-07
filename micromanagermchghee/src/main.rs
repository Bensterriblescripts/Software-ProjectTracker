use windows::{
    core::*,
    core::GUID,
    Win32::System::Com::*,
    Win32::UI::Shell::*,
    Win32::Foundation::HWND,
    Win32::UI::WindowsAndMessaging::GetForegroundWindow
};
use std::{
    io::{self, Write},
    thread::sleep,
    time::Duration,
};
use windows::Win32::Foundation::GetLastError;

struct VirtualDesktopManager {
    manager: IVirtualDesktopManager,
}
const VIRTUAL_DESKTOP_MANAGER_CLSID: GUID = GUID::from_u128(0xaa509086_5ca9_4c25_8f95_589d3c07b48a);
impl VirtualDesktopManager {
    fn new() -> Result<Self> {
        let manager: IVirtualDesktopManager = unsafe {
            CoCreateInstance(
                &VIRTUAL_DESKTOP_MANAGER_CLSID,
                None,
                CLSCTX_INPROC_SERVER,
            )?
        };

        Ok(Self { manager })
    }

    // fn is_window_on_current_desktop(&self, hwnd: HWND) -> Result<bool> {
    //     unsafe { Ok(self.manager.IsWindowOnCurrentVirtualDesktop(hwnd)?.as_bool()) }
    // }

    fn get_window_desktop_id(&self, hwnd: HWND) -> Result<GUID> {
        unsafe { self.manager.GetWindowDesktopId(hwnd) }
    }
}
fn getdesktopid() -> Option<GUID> {

    // Retrieve
    let result = (|| {
        let vdm = VirtualDesktopManager::new().ok()?;
        let hwnd = unsafe { GetForegroundWindow() };
        vdm.get_window_desktop_id(hwnd).ok()
    })();

    result
}
fn main() {

    // Init Memory
    let mut desktops = std::collections::HashMap::with_capacity(8);
    if unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).is_err() } {
        println!("CoInitializeEx failed with error: {:?}", unsafe { GetLastError() });
    };

    // Core
    loop {

        // Retrieve ID
        let Some(current_desktop_id) = getdesktopid() else {
            eprintln!("Failed to get desktop ID");
            return;
        };

        // Desktop Change
        if !desktops.contains_key(&current_desktop_id) {
            print!("\nNew Desktop ID. Enter an alias: ");
            io::stdout().flush().unwrap();

            let mut input = String::with_capacity(32);
            if io::stdin().read_line(&mut input).is_ok() {
                desktops.insert(current_desktop_id, input.trim().to_string());
            }
            println!("Created New Desktop: {}", input);
            println!("Current Desktops:");
             for (key, value) in desktops.iter() {
                println!("{key:?}: {value}");
            }
        }

        sleep(Duration::from_millis(100));
    }

    unsafe { CoUninitialize() };

}