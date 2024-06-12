use std::{env, fs};
use std::fs::File;
use std::io::Write;
use std::path::{Path};
use std::process::{Command, Stdio};
use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
use winapi::shared::minwindef::DWORD;
use winapi::um::errhandlingapi::GetLastError;
use winapi::um::winuser::{SystemParametersInfoW, SPI_SETDESKWALLPAPER, SPIF_UPDATEINIFILE, SPIF_SENDWININICHANGE};

//为了编译成单个可执行exe文件，将所需文件使用include_bytes!宏导入编译进程序
//壁纸图片两张
pub const WALLPAPER_DATA: &[u8] = include_bytes!("./wallpaper.jpg");
pub const LOCK_SCREEN_DATA: &[u8] = include_bytes!("./lockscreen.jpg");

//设置windows锁屏壁纸所需的第三方开源工具 ImageGlass/igcmd.exe
pub const IGCMD_DLL_DATA: &[u8] = include_bytes!("./ImageGlass/igcmd.dll");
pub const IGCMD_DLL_CFG_DATA: &[u8] = include_bytes!("./ImageGlass/igcmd.dll.config");
pub const IGCMD_EXE_DATA: &[u8] = include_bytes!("./ImageGlass/igcmd.exe");
pub const IGCMD_PDB_DATA: &[u8] = include_bytes!("./ImageGlass/igcmd.pdb");
pub const IGCMD_RUNTIMECFG_DATA: &[u8] = include_bytes!("./ImageGlass/igcmd.runtimeconfig.json");
pub const IMG_BASEDLL_DATA: &[u8] = include_bytes!("./ImageGlass/ImageGlass.Base.dll");
pub const IMG_RTMCFGJSON_DATA: &[u8] = include_bytes!("./ImageGlass/ImageGlass.runtimeconfig.json");
pub const IMG_SETDLL_DATA: &[u8] = include_bytes!("./ImageGlass/ImageGlass.Settings.dll");
pub const IMG_SETPDB_DATA: &[u8] = include_bytes!("./ImageGlass/ImageGlass.Settings.pdb");
pub const IMG_UIDLL_DATA: &[u8] = include_bytes!("./ImageGlass/ImageGlass.UI.dll");
pub const MS_ESTCFGCMDLINE_DATA: &[u8] = include_bytes!("./ImageGlass/Microsoft.Extensions.Configuration.CommandLine.dll");
pub const MS_ESTABS_DATA: &[u8] = include_bytes!("./ImageGlass/Microsoft.Extensions.FileProviders.Abstractions.dll");
pub const MS_PHYSICAL_DATA: &[u8] = include_bytes!("./ImageGlass/Microsoft.Extensions.FileProviders.Physical.dll");
pub const MS_FSGLOB_DATA: &[u8] = include_bytes!("./ImageGlass/Microsoft.Extensions.FileSystemGlobbing.dll");
pub const MS_PRIMI_DATA: &[u8] = include_bytes!("./ImageGlass/Microsoft.Extensions.Primitives.dll");
pub const MS_SDK_DATA: &[u8] = include_bytes!("./ImageGlass/Microsoft.Windows.SDK.NET.dll");
pub const WINRT_DATA: &[u8] = include_bytes!("./ImageGlass/WinRT.Runtime.dll");
pub const MS_ABS_DATA: &[u8] = include_bytes!("./ImageGlass/Microsoft.Extensions.Configuration.Abstractions.dll");

// igcmd.exe运行需要.NET 8.0 Desktop Runtime也编译进exe文件,运行时解压缩出来并且使用命令行静默安装
pub const WINDOWS_DESKTOP_RUNTIME: &[u8] = include_bytes!("./windowsdesktop-runtime-8.0.6-win-x64.exe");


fn save_to_file(path: &Path, data: &[u8]) -> std::io::Result<()> {
    let mut file = File::create(&path)?;
    file.write_all(data)?;
    Ok(())
}

fn set_desktop_wallpaper(path: &str) -> Result<(), DWORD> {
    let widestring: Vec<u16> = OsStr::new(path).encode_wide().chain(Some(0).into_iter()).collect();
    let result = unsafe {
        SystemParametersInfoW(
            SPI_SETDESKWALLPAPER,
            0,
            widestring.as_ptr() as *mut _,
            SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE,
        )
    };

    if result != 0 {
        Ok(())
    } else {
        Err(unsafe { GetLastError() })
    }
}

fn set_lock_screen_wallpaper(igcmd_path: &Path, image_path: &Path) -> Result<(), std::io::Error> {
    let output = Command::new(igcmd_path)
        .arg("set-lock-screen")
        .arg(image_path)
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .output()?;

    if output.status.success() {
        Ok(())
    } else {
        Err(std::io::Error::new(std::io::ErrorKind::Other, "Failed to set lock screen wallpaper"))
    }
}

// 安装 windowsdesktop-runtime-8.0.6-win-x64.exe
fn install_windesktop_runtime(path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    // 运行安装程序
    let status = Command::new(path)
        .arg("/install")
        .arg("/quiet")
        .arg("/norestart")
        .status()?;

    if !status.success() {
        return Err(format!("安装程序退出码: {}", status).into());
    }

    Ok(())
}

fn main() {
    // Get current path
    let current_exe_path = env::current_exe().expect("Failed to get current exe path");
    let current_dir = current_exe_path.parent().expect("Failed to get current directory");
    println!("{:?}", current_exe_path);
    println!("{:?}", current_dir);

    let wallpaper_path = current_dir.join("wallpaper.jpg");
    let lock_screen_path = current_dir.join("lockscreen.jpg");
    let imgglass_path = current_dir.join("ImageGlass");

    // 将.NET 8.0 Desktop Runtime安装文件写入当前路径
    let win_desktop_runtime_path = current_dir.join("windowsdesktop-runtime-8.0.6-win-x64.exe");
    save_to_file(&win_desktop_runtime_path, WINDOWS_DESKTOP_RUNTIME).unwrap();

    println!("开始安装 .NET 8.0 Desktop Runtime (v8.0.6)...");
    match install_windesktop_runtime(&win_desktop_runtime_path) {
        Ok(_) => println!("安装.NET 8.0 Desktop Runtime (v8.0.6) 完成!"),
        Err(_e) => println!("安装.NET Desktop Runtime 失败！"),
    }

    // 创建ImageGlass文件夹
    fs::create_dir_all(&imgglass_path).expect("TODO: 创建文件夹异常");

    // Save image data to files
    save_to_file(&wallpaper_path, WALLPAPER_DATA).unwrap();
    save_to_file(&lock_screen_path, LOCK_SCREEN_DATA).unwrap();

    println!("图片已写入本地！");
    // println!("{:?}", imgglass_path);
    // 保存ImageGlass相关文件到当前目录
    let igcmd_dll_path = imgglass_path.join("igcmd.dll");
    save_to_file(&igcmd_dll_path, IGCMD_DLL_DATA).unwrap();
    let igcmd_dllcfg_path = imgglass_path.join("jgcmd.dll.config");
    save_to_file(&igcmd_dllcfg_path, IGCMD_DLL_CFG_DATA).unwrap();
    let igcmd_exe_path = imgglass_path.join("igcmd.exe");
    save_to_file(&igcmd_exe_path, IGCMD_EXE_DATA).unwrap();
    let igcmd_pdb_path = imgglass_path.join("igcmd.pdb");
    save_to_file(&igcmd_pdb_path, IGCMD_PDB_DATA).unwrap();
    let igcmd_runtimecfg_path = imgglass_path.join("igcmd.runtimeconfig.json");
    save_to_file(&igcmd_runtimecfg_path, IGCMD_RUNTIMECFG_DATA).unwrap();
    let img_basedll_path = imgglass_path.join("ImageGlass.Base.dll");
    save_to_file(&img_basedll_path, IMG_BASEDLL_DATA).unwrap();
    let img_rtmcfgjson_path = imgglass_path.join("ImageGlass.runtimeconfig.json");
    save_to_file(&img_rtmcfgjson_path, IMG_RTMCFGJSON_DATA).unwrap();
    let img_setdll_path = imgglass_path.join("ImageGlass.Settings.dll");
    save_to_file(&img_setdll_path, IMG_SETDLL_DATA).unwrap();
    let img_setpdb_path = imgglass_path.join("ImageGlass.Settings.pdb");
    save_to_file(&img_setpdb_path, IMG_SETPDB_DATA).unwrap();
    let img_uidll_path = imgglass_path.join("ImageGlass.UI.dll");
    save_to_file(&img_uidll_path, IMG_UIDLL_DATA).unwrap();
    let ms_estcfgcmdline_path = imgglass_path.join("Microsoft.Extensions.Configuration.CommandLine.dll");
    save_to_file(&ms_estcfgcmdline_path, MS_ESTCFGCMDLINE_DATA).unwrap();
    let ms_estabs_path = imgglass_path.join("Microsoft.Extensions.FileProviders.Abstractions.dll");
    save_to_file(&ms_estabs_path, MS_ESTABS_DATA).unwrap();
    let ms_physical_path = imgglass_path.join("Microsoft.Extensions.FileProviders.Physical.dll");
    save_to_file(&ms_physical_path, MS_PHYSICAL_DATA).unwrap();
    let ms_fsglob_path = imgglass_path.join("Microsoft.Extensions.FileSystemGlobbing.dll");
    save_to_file(&ms_fsglob_path, MS_FSGLOB_DATA).unwrap();
    let ms_primi_path = imgglass_path.join("Microsoft.Extensions.Primitives.dll");
    save_to_file(&ms_primi_path, MS_PRIMI_DATA).unwrap();
    let ms_sdk_path = imgglass_path.join("Microsoft.Windows.SDK.NET.dll");
    save_to_file(&ms_sdk_path, MS_SDK_DATA).unwrap();
    let winrt_path = imgglass_path.join("WinRT.Runtime.dll");
    save_to_file(&winrt_path, WINRT_DATA).unwrap();
    let ms_abs_path = imgglass_path.join("Microsoft.Extensions.Configuration.Abstractions.dll");
    save_to_file(&ms_abs_path, MS_ABS_DATA).unwrap();

    // Set desktop wallpaper
    match set_desktop_wallpaper(wallpaper_path.to_str().unwrap()) {
        Ok(_) => println!("Desktop wallpaper set successfully!"),
        Err(e) => println!("Failed to set desktop wallpaper. Error code: {}", e),
    }

    // Set lock screen wallpaper
    match set_lock_screen_wallpaper(&igcmd_exe_path, &lock_screen_path) {
        Ok(_) => println!("Lock screen wallpaper set successfully!"),
        Err(_e) => println!("Failed to set lock screen wallpaper."),
    }

    // 设置完成 删除ImageGlass文件夹,删除 windowsdesktop-runtime-8.0.6-win-x64.exe, 删除图片
    fs::remove_dir_all(imgglass_path).expect("TODO: 删除ImageGlass文件夹异常！");
    fs::remove_file(win_desktop_runtime_path).expect("TODO: 删除windowsdesktop-runtime-8.0.6-win-x64.exe异常！");
    fs::remove_file(wallpaper_path).expect("TODO: 删除 wallpaper.jpg 异常！");
    fs::remove_file(lock_screen_path).expect("TODO: 删除 lockscreen.jpg 异常！");
}
