import argparse
import os
from fractions import Fraction
from typing import Any, Dict, Iterable, Optional, Tuple

import piexif
from PIL import ExifTags, Image


GPS_TAGS = ExifTags.GPSTAGS
EXIF_TAGS = ExifTags.TAGS


def format_value(value: Any) -> str:
    if value is None:
        return "N/A"
    if isinstance(value, bytes):
        try:
            return value.decode("utf-8", errors="replace")
        except Exception:
            return str(value)
    if isinstance(value, tuple):
        return ", ".join(format_value(v) for v in value)
    return str(value)


def rational_to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    try:
        if isinstance(value, tuple) and len(value) == 2 and value[1] != 0:
            return float(value[0]) / float(value[1])
        return float(value)
    except Exception:
        return None


def fraction_string(value: Any) -> str:
    if isinstance(value, tuple) and len(value) == 2:
        return f"{value[0]}/{value[1]}"
    if isinstance(value, (int, float)):
        frac = Fraction(value).limit_denominator(1000000)
        return f"{frac.numerator}/{frac.denominator}"
    return format_value(value)


def detect_format(image_path: str) -> Tuple[str, str]:
    ext = os.path.splitext(image_path)[1].lower()
    with open(image_path, "rb") as f:
        header = f.read(16)
    if header.startswith(b"\xFF\xD8\xFF"):
        detected = "jpeg"
    elif len(header) >= 12 and header[4:8] == b"ftyp" and any(
        marker in header for marker in (b"heic", b"heix", b"hevc", b"hevx")
    ):
        detected = "heic"
    else:
        detected = "unknown"
    return ext, header.hex()


def load_image(image_path: str):
    ext = os.path.splitext(image_path)[1].lower()
    if ext == ".heic":
        try:
            from pillow_heif import register_heif_opener

            register_heif_opener()
        except Exception as err:
            raise RuntimeError(
                "HEIC support failed. Suggestion: install libheif and reinstall pillow-heif, "
                "or convert first using: heif-convert photo.heic photo.jpg"
            ) from err
    return Image.open(image_path)


def to_named_exif(exif_obj) -> Dict[str, Any]:
    named: Dict[str, Any] = {}
    if not exif_obj:
        return named
    for tag, value in exif_obj.items():
        name = EXIF_TAGS.get(tag, str(tag))
        named[name] = value
    return named


def get_gps_named(gps_ifd: Dict[int, Any]) -> Dict[str, Any]:
    return {GPS_TAGS.get(k, str(k)): v for k, v in gps_ifd.items()}


def dms_to_decimal(dms: Any, ref: str) -> Optional[float]:
    if not dms or len(dms) != 3:
        return None
    deg = rational_to_float(dms[0])
    mins = rational_to_float(dms[1])
    secs = rational_to_float(dms[2])
    if deg is None or mins is None or secs is None:
        return None
    dec = deg + (mins / 60.0) + (secs / 3600.0)
    if ref in ("S", "W"):
        dec *= -1
    return dec


def map_lookup(mapping: Dict[int, str], key: Any) -> str:
    if key is None:
        return "N/A"
    return mapping.get(key, str(key))


def markdown_table(rows: Iterable[Tuple[str, str]]) -> str:
    lines = ["| Field | Value |", "|---|---|"]
    for field, value in rows:
        safe_field = field.replace("|", "\\|")
        safe_value = value.replace("|", "\\|")
        lines.append(f"| {safe_field} | {safe_value} |")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="Extract EXIF metadata from JPEG/HEIC image.")
    parser.add_argument("image_path", help="Path to image file (.jpg, .jpeg, .heic)")
    args = parser.parse_args()

    image_path = args.image_path
    if not os.path.exists(image_path):
        print(f"File not found: {image_path}")
        return

    ext, magic_hex = detect_format(image_path)
    if ext not in (".jpg", ".jpeg", ".heic"):
        print("Unsupported format for this EXIF workflow.")
        print("GPS EXIF is typically not present in PNG/other formats.")
        print("Suggestion: re-save as JPEG from the original camera app.")
        return

    try:
        img = load_image(image_path)
    except RuntimeError as heic_error:
        print(str(heic_error))
        print("Linux fallback:")
        print("  sudo apt install libheif-dev")
        print("  pip install --force-reinstall pillow-heif")
        print("or convert first:")
        print("  heif-convert photo.heic photo.jpg")
        return

    raw_exif = img.info.get("exif")
    exif_obj = img.getexif()
    named_exif = to_named_exif(exif_obj)
    piexif_data = piexif.load(raw_exif) if raw_exif else None

    if not named_exif and not piexif_data:
        print("No EXIF metadata found in this file")
        return

    zeroth = piexif_data.get("0th", {}) if piexif_data else {}
    exif_ifd = piexif_data.get("Exif", {}) if piexif_data else {}
    gps_ifd = piexif_data.get("GPS", {}) if piexif_data else {}
    gps_named = get_gps_named(gps_ifd)

    # Enum mappings for friendlier values
    white_balance_map = {0: "Auto", 1: "Manual"}
    metering_mode_map = {
        0: "Unknown",
        1: "Average",
        2: "Center-weighted",
        3: "Spot",
        4: "Multi-spot",
        5: "Pattern",
        6: "Partial",
        255: "Other",
    }
    exposure_program_map = {
        0: "Not defined",
        1: "Manual",
        2: "Normal program",
        3: "Aperture priority",
        4: "Shutter priority",
        5: "Creative program",
        6: "Action program",
        7: "Portrait mode",
        8: "Landscape mode",
    }
    exposure_mode_map = {0: "Auto exposure", 1: "Manual exposure", 2: "Auto bracket"}
    scene_capture_map = {
        0: "Standard",
        1: "Landscape",
        2: "Portrait",
        3: "Night scene",
    }
    flash_map = {
        0: "Flash did not fire",
        1: "Flash fired",
        5: "Flash fired, strobe return light not detected",
        7: "Flash fired, strobe return light detected",
        9: "Flash fired, compulsory flash mode",
        16: "Flash did not fire, compulsory flash mode",
        24: "Flash did not fire, auto mode",
        25: "Flash fired, auto mode",
        32: "No flash function",
    }
    color_space_map = {1: "sRGB", 65535: "Uncalibrated"}

    make = format_value(zeroth.get(piexif.ImageIFD.Make) or named_exif.get("Make"))
    model = format_value(zeroth.get(piexif.ImageIFD.Model) or named_exif.get("Model"))
    lens = format_value(exif_ifd.get(piexif.ExifIFD.LensModel) or named_exif.get("LensModel"))
    dt = format_value(
        exif_ifd.get(piexif.ExifIFD.DateTimeOriginal)
        or named_exif.get("DateTimeOriginal")
        or named_exif.get("DateTime")
    )

    exposure_time = exif_ifd.get(piexif.ExifIFD.ExposureTime) or named_exif.get("ExposureTime")
    fnumber = exif_ifd.get(piexif.ExifIFD.FNumber) or named_exif.get("FNumber")
    iso = exif_ifd.get(piexif.ExifIFD.ISOSpeedRatings) or named_exif.get("ISOSpeedRatings")
    focal = exif_ifd.get(piexif.ExifIFD.FocalLength) or named_exif.get("FocalLength")
    focal35 = exif_ifd.get(piexif.ExifIFD.FocalLengthIn35mmFilm) or named_exif.get("FocalLengthIn35mmFilm")
    flash = exif_ifd.get(piexif.ExifIFD.Flash) or named_exif.get("Flash")
    wb = exif_ifd.get(piexif.ExifIFD.WhiteBalance) or named_exif.get("WhiteBalance")
    metering = exif_ifd.get(piexif.ExifIFD.MeteringMode) or named_exif.get("MeteringMode")
    exposure_program = exif_ifd.get(piexif.ExifIFD.ExposureProgram) or named_exif.get("ExposureProgram")
    exposure_mode = exif_ifd.get(piexif.ExifIFD.ExposureMode) or named_exif.get("ExposureMode")
    scene_capture = exif_ifd.get(piexif.ExifIFD.SceneCaptureType) or named_exif.get("SceneCaptureType")

    width = img.width
    height = img.height
    dpi = img.info.get("dpi")
    orientation = zeroth.get(piexif.ImageIFD.Orientation) or named_exif.get("Orientation")
    color_space = exif_ifd.get(piexif.ExifIFD.ColorSpace) or named_exif.get("ColorSpace")
    sensing_method = exif_ifd.get(piexif.ExifIFD.SensingMethod) or named_exif.get("SensingMethod")

    lat_dms = gps_named.get("GPSLatitude")
    lat_ref = gps_named.get("GPSLatitudeRef")
    lon_dms = gps_named.get("GPSLongitude")
    lon_ref = gps_named.get("GPSLongitudeRef")
    altitude = gps_named.get("GPSAltitude")
    gps_time = gps_named.get("GPSTimeStamp")
    gps_date = gps_named.get("GPSDateStamp")

    lat_ref_text = format_value(lat_ref)
    lon_ref_text = format_value(lon_ref)
    lat_decimal = dms_to_decimal(lat_dms, lat_ref_text) if lat_dms and lat_ref else None
    lon_decimal = dms_to_decimal(lon_dms, lon_ref_text) if lon_dms and lon_ref else None

    print("# EXIF Metadata Report")
    print()
    print("## Detection")
    detection_rows = [
        ("File", image_path),
        ("Extension", ext),
        ("Magic Bytes (first 16 bytes)", magic_hex),
        ("Detected Format", "HEIC" if ext == ".heic" else "JPEG"),
    ]
    print(markdown_table((k, format_value(v)) for k, v in detection_rows))
    print()

    print("## Camera Info")
    camera_rows = [
        ("Make", make),
        ("Model", model),
        ("Lens Model", lens),
    ]
    print(markdown_table(camera_rows))
    print()

    print("## Camera Settings")
    camera_settings_rows = [
        ("Date/Time Taken", dt),
        ("Exposure Time", f"{fraction_string(exposure_time)} sec" if exposure_time else "N/A"),
        ("Aperture (f-number)", f"f/{rational_to_float(fnumber):.2f}" if rational_to_float(fnumber) else "N/A"),
        ("ISO", format_value(iso)),
        ("Focal Length", f"{rational_to_float(focal):.2f} mm" if rational_to_float(focal) else "N/A"),
        ("Focal Length (35mm equivalent)", f"{format_value(focal35)} mm" if focal35 else "N/A"),
        ("Flash Status", map_lookup(flash_map, flash)),
        ("White Balance", map_lookup(white_balance_map, wb)),
        ("Metering Mode", map_lookup(metering_mode_map, metering)),
        ("Exposure Program", map_lookup(exposure_program_map, exposure_program)),
        ("Exposure Mode", map_lookup(exposure_mode_map, exposure_mode)),
        ("Scene Capture Type", map_lookup(scene_capture_map, scene_capture)),
    ]
    print(markdown_table(camera_settings_rows))
    print()

    print("## Image Info")
    dpi_str = "N/A"
    if isinstance(dpi, tuple) and len(dpi) >= 2:
        dpi_str = f"{dpi[0]} x {dpi[1]} DPI"
    image_rows = [
        ("Width", f"{width} px"),
        ("Height", f"{height} px"),
        ("Resolution (DPI)", dpi_str),
        ("Orientation", format_value(orientation)),
        ("Color Space", map_lookup(color_space_map, color_space)),
        ("Sensing Method", format_value(sensing_method)),
    ]
    print(markdown_table(image_rows))
    print()

    print("## GPS Info")
    if not gps_ifd:
        print("No GPS data embedded in this photo")
    else:
        altitude_m = (
            f"{rational_to_float(altitude):.2f} meters" if rational_to_float(altitude) is not None else format_value(altitude)
        )
        gps_rows = [
            ("GPSLatitude (raw DMS rational)", format_value(lat_dms)),
            ("GPSLatitudeRef", lat_ref_text),
            ("GPSLongitude (raw DMS rational)", format_value(lon_dms)),
            ("GPSLongitudeRef", lon_ref_text),
            ("GPS Altitude", altitude_m),
            ("GPS TimeStamp", format_value(gps_time)),
            ("GPS Date", format_value(gps_date)),
            ("Latitude Decimal", f"{lat_decimal:.8f}" if lat_decimal is not None else "N/A"),
            ("Longitude Decimal", f"{lon_decimal:.8f}" if lon_decimal is not None else "N/A"),
        ]
        print(markdown_table(gps_rows))
        print()

        if lat_decimal is not None and lon_decimal is not None:
            print("### Google Maps Link")
            print(f"https://maps.google.com/?q={lat_decimal:.8f},{lon_decimal:.8f}")
            print()

        print("### EXIF-ready GPS Rational Format")
        if lat_dms and lon_dms:
            lat0 = fraction_string(lat_dms[0])
            lat1 = fraction_string(lat_dms[1])
            lat2 = fraction_string(lat_dms[2])
            lon0 = fraction_string(lon_dms[0])
            lon1 = fraction_string(lon_dms[1])
            lon2 = fraction_string(lon_dms[2])
            print(f"GPSLatitude:     {lat0}, {lat1}, {lat2}")
            print(f"GPSLatitudeRef:  {lat_ref_text}")
            print(f"GPSLongitude:    {lon0}, {lon1}, {lon2}")
            print(f"GPSLongitudeRef: {lon_ref_text}")
        else:
            print("GPS rational coordinate values are incomplete.")
    print()

    print("## Notes / Edge Cases")
    print("- If this image came from social media (Facebook/Instagram/Viber/WhatsApp), metadata is often stripped and cannot be recovered.")
    print("- If this file is a screenshot, camera EXIF usually does not exist.")
    print("- South/West references (S/W) are negated in decimal conversion.")
    print("- Edited exports (Snapseed/Lightroom) may preserve or strip GPS selectively.")


if __name__ == "__main__":
    main()
