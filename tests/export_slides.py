import os
import sys
try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

def export_presentation(pptx_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    if not WIN32_AVAILABLE:
        print("[INFO] win32com not available (non-Windows/headless). Skipping PowerPoint COM export.")
        return True
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    except Exception as e:
        print(f"[INFO] PowerPoint COM automation not available ({e}). Skipping slide export.")
        return True
    try:
        # Open in read-only, hidden / minimized
        pres = ppt_app.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
        for i, slide in enumerate(pres.Slides, start=1):
            out_img = os.path.abspath(os.path.join(output_dir, f"slide_{i}.png"))
            slide.Export(out_img, "PNG", 1920, 1080)
            print(f"Exported Slide {i} -> {out_img}")
        pres.Close()
        return True
    finally:
        try:
            ppt_app.Quit()
        except Exception:
            pass

if __name__ == "__main__":
    pptx = sys.argv[1] if len(sys.argv) > 1 else r"."
    out = sys.argv[2] if len(sys.argv) > 2 else r"."
    export_presentation(pptx, out)
