use windows::{
    core::*,
    core::GUID,
    Win32::{
        System::Com::*,
        UI::Shell::*,
        Foundation::HWND,
        UI::WindowsAndMessaging::GetForegroundWindow,
        Foundation::GetLastError,
    }
};
use std::{
    io::{self, Write},
    time::Duration,
    env,
    fs,
    sync::{Arc, Mutex},
    thread,
    thread::sleep
};


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
    let desktops = Arc::new(Mutex::new(std::collections::HashMap::with_capacity(8)));
    if unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).is_err() } {
        println!("CoInitializeEx failed with error: {:?}", unsafe { GetLastError() });
    };

    // Init Window
    let testing = "Independent Assessor";
    let timer = Timer::new().unwrap();
    timer.set_display_text(testing.into());

    let window = timer.window();
    window.set_position(slint::PhysicalPosition::new(1250, 1025)); // Window Position

    let weak = timer.as_weak(); // Weak Pointer
    let desktops_clone = Arc::clone(&desktops); // Clone Desktop

    // Core Functionality
    thread::spawn(move || {
        loop {
            if let Some(current_desktop_id) = getdesktopid() {
                let mut desktops = desktops_clone.lock().unwrap();

                if !desktops.contains_key(&current_desktop_id) {
                    // Update UI from background thread
                    weak.upgrade_in_event_loop(move |handle| {
                        handle.set_display_text("New Desktop".into());
                    }).unwrap();

                    print!("\nNew Desktop ID. Enter an alias: ");
                    io::stdout().flush().unwrap();

                    let mut input = String::with_capacity(32);
                    if io::stdin().read_line(&mut input).is_ok() {
                        let input_trimmed = input.trim().to_string();
                        desktops.insert(current_desktop_id, input_trimmed.clone());

                        // Update UI with new desktop name
                        weak.upgrade_in_event_loop(move |handle| {
                            handle.set_display_text(slint::SharedString::from(input_trimmed));
                        }).unwrap();
                    }

                    println!("Current Desktops:");
                    for (key, value) in desktops.iter() {
                        println!("{key:?}: {value}");
                    }
                }
                else if desktops.contains_key(&current_desktop_id) {
                    weak.upgrade_in_event_loop(move |handle| {
                        handle.set_display_text(slint::SharedString::from(desktops.get(&current_desktop_id).unwrap()));
                    }).unwrap();
                }
            }

            sleep(Duration::from_millis(100));
        }
    });

    timer.run().unwrap(); // Start UI
}