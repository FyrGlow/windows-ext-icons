use image::{ImageBuffer, RgbaImage};
use windows::core::PCWSTR;
use windows::Win32::{
    Graphics::Gdi::{CreateCompatibleDC, DeleteDC, DeleteObject, GetDIBits, SelectObject, BITMAPINFO, BITMAPINFOHEADER, DIB_RGB_COLORS},
    Storage::FileSystem::FILE_FLAGS_AND_ATTRIBUTES,
    UI::{
        Controls::{IImageList, ILD_TRANSPARENT},
        Shell::{SHGetFileInfoW, SHGetImageList, SHFILEINFOW, SHGFI_SYSICONINDEX, SHIL_SMALL, SHIL_LARGE, SHIL_EXTRALARGE, SHIL_JUMBO},
        WindowsAndMessaging::{DestroyIcon, GetIconInfoExW, HICON, ICONINFOEXW},
    },
};

use std::arch::x86_64::{__m128i, _mm_loadu_si128, _mm_setr_epi8, _mm_shuffle_epi8, _mm_storeu_si128};
use std::os::windows::ffi::OsStrExt;
use std::path::Path;

/// Fetches an icon as an image from a given file path and specified icon size flag.
pub fn fetch_icon_as_image(
    path: &Path, 
    icon_size_flag: i32
) -> Result<RgbaImage, Box<dyn std::error::Error>> {
    unsafe {
        let wide_path: Vec<u16> = path.as_os_str().encode_wide().chain(Some(0)).collect();
        let mut file_info = SHFILEINFOW::default();

        if SHGetFileInfoW(
            PCWSTR(wide_path.as_ptr()),
            FILE_FLAGS_AND_ATTRIBUTES(0),
            Some(&mut file_info),
            std::mem::size_of::<SHFILEINFOW>() as u32,
            SHGFI_SYSICONINDEX,
        ) == 0 || file_info.iIcon == 0
        {
            return Err("Failed to retrieve icon".into());
        }

        let image_list: IImageList = SHGetImageList(icon_size_flag)?;
        let icon = image_list.GetIcon(file_info.iIcon, ILD_TRANSPARENT.0)?;
        let image = hicon_to_image(&icon)?;

        DestroyIcon(icon)?;
        Ok(image)
    }
}

/// Converts a handle to an icon (HICON) into an image buffer (RgbaImage).
pub fn hicon_to_image(hicon: &HICON) -> Result<RgbaImage, Box<dyn std::error::Error>> {
    unsafe {
        let mut icon_info = ICONINFOEXW {
            cbSize: std::mem::size_of::<ICONINFOEXW>() as u32,
            ..Default::default()
        };

        if !GetIconInfoExW(*hicon, &mut icon_info).as_bool() {
            return Err("Failed to retrieve icon information".into());
        }

        let screen_dc = CreateCompatibleDC(None);
        let mem_dc = CreateCompatibleDC(screen_dc);
        let old_bitmap = SelectObject(mem_dc, icon_info.hbmColor);

        let mut bmp_info = BITMAPINFO {
            bmiHeader: BITMAPINFOHEADER {
                biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                biWidth: icon_info.xHotspot as i32 * 2,
                biHeight: -(icon_info.yHotspot as i32 * 2),
                biPlanes: 1,
                biBitCount: 32,
                biCompression: DIB_RGB_COLORS.0,
                ..Default::default()
            },
            ..Default::default()
        };

        let mut pixel_data = vec![0; (icon_info.xHotspot * 2 * icon_info.yHotspot * 2 * 4) as usize];

        if GetDIBits(
            mem_dc,
            icon_info.hbmColor,
            0,
            icon_info.yHotspot * 2,
            Some(pixel_data.as_mut_ptr() as *mut _),
            &mut bmp_info,
            DIB_RGB_COLORS,
        ) == 0 {
            return Err("Failed to retrieve bitmap data".into());
        }

        SelectObject(mem_dc, old_bitmap);
        DeleteDC(mem_dc).ok()?;
        DeleteDC(screen_dc).ok()?;
        DeleteObject(icon_info.hbmColor).ok()?;
        DeleteObject(icon_info.hbmMask).ok()?;

        if bmp_info.bmiHeader.biBitCount != 32 {
            return Err("Icon is not 32-bit".into());
        }

        bgra_to_rgba(&mut pixel_data);
        let image = ImageBuffer::from_raw(
            icon_info.xHotspot * 2,
            icon_info.yHotspot * 2,
            pixel_data,
        ).expect("Failed to create image buffer");

        Ok(image)
    }
}

/// Converts pixel data from BGRA format to RGBA format in place.
pub fn bgra_to_rgba(data: &mut [u8]) {
    let mask: __m128i = unsafe {
        _mm_setr_epi8(
            2, 1, 0, 3,
            6, 5, 4, 7,
            10, 9, 8, 11,
            14, 13, 12, 15,
        )
    };

    for chunk in data.chunks_exact_mut(16) {
        let vector = unsafe { _mm_loadu_si128(chunk.as_ptr() as *const __m128i) };
        let shuffled = unsafe { _mm_shuffle_epi8(vector, mask) };
        unsafe { _mm_storeu_si128(chunk.as_mut_ptr() as *mut __m128i, shuffled) };
    }
}


