use windows::{
    core::*,
    core::GUID,
    Win32::System::Com::*,
    Win32::UI::Shell::*,
    Win32::Foundation::HWND,
    Win32::Foundation::HANDLE,
    Win32::UI::WindowsAndMessaging::GetForegroundWindow
};
use std::{
    io::{self, Write},
    thread::sleep,
    time::Duration,
    env,
    fs
};
use windows::Win32::Foundation::GetLastError;

/* UI */
slint::slint!{
    export component Timer inherits Window {
        in property <string> display_text;

        preferred-width: 300px;
        preferred-height: 100px;
        title: "Micromanager Mcghee";
        no-frame: true;
        background: transparent;
        always-on-top: true;

        Text {
            text: root.display_text;
            color: white;
            font-family: "Calibri";
            font-size: 24px;
        }
    }
}

/* File System */
fn check_program_folder() -> Result<()> {
    let app_folder = format!("{}\\Micromanager Mcghee", env::var("LOCALAPPDATA").unwrap_or_default());
    Ok(fs::create_dir_all(app_folder)?)
}

/* Virtual Desktops */
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

    fn is_window_on_current_desktop(&self, hwnd: HWND) -> Result<bool> {
        unsafe { Ok(self.manager.IsWindowOnCurrentVirtualDesktop(hwnd)?.as_bool()) }
    }

    fn get_window_desktop_id(&self, hwnd: HWND) -> Result<GUID> {
        unsafe { self.manager.GetWindowDesktopId(hwnd) }
    }
}
fn getdesktopid() -> Option<GUID> {
    let result = (|| {
        let vdm = VirtualDesktopManager::new().ok()?;
        let hwnd = unsafe { GetForegroundWindow() };

        // Make sure the window is on the new desktop
        vdm.is_window_on_current_desktop(hwnd).ok();
        vdm.get_window_desktop_id(hwnd).ok()
    })();

    result
}

fn main() {

    // Init Storage
    if check_program_folder().is_err() {
        println!("Error creating folder");
        return;
    };

    // Init Memory
    let mut desktops = std::collections::HashMap::with_capacity(8);
    if unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).is_err() } {
        println!("CoInitializeEx failed with error: {:?}", unsafe { GetLastError() });
    };

    // Init Timer
    let testing = "Independent Assessor";
    let timer = Timer::new().unwrap();
    timer.set_display_text(testing.into());

    let window = timer.window();
    window.set_position(slint::PhysicalPosition::new(1600,0)); // Window Positiom

    let weak = timer.as_weak();

    timer.run().unwrap(); // Start

    // Core
    loop {

        let Some(current_desktop_id) = getdesktopid() else {
            continue;
        };

        if !desktops.contains_key(&current_desktop_id) {

            weak.upgrade_in_event_loop(move |handle| {
                handle.set_display_text("New Desktop".into());
            }).unwrap();

            print!("\nNew Desktop ID. Enter an alias: ");
            io::stdout().flush().unwrap();

            let mut input = String::with_capacity(32); // Get New Alias
            if io::stdin().read_line(&mut input).is_ok() {
                desktops.insert(current_desktop_id, input.trim().to_string());
            }
            weak.upgrade_in_event_loop(move |handle| { // Update Text
                handle.set_display_text(input.into());
            }).unwrap();
            println!("Current Desktops:");
             for (key, value) in desktops.iter() {
                println!("{key:?}: {value}");
            }
        }

        sleep(Duration::from_millis(100));
    }

    unsafe { CoUninitialize() };
}