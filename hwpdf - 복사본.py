import os
import win32com.client

FOLDER = r"D:\NaverCloud\lig\수학\프라하작업장\내신대비자료\2026\2026 중3\2026 중3 1학기 1차 대비"


def main():
    hwp_files = [f for f in os.listdir(FOLDER) if f.lower().endswith(".hwp")]
    if not hwp_files:
        print("HWP 파일이 없습니다.")
        return

    hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    hwp.XHwpWindows.Item(0).Visible = False

    success, fail = 0, 0
    for filename in hwp_files:
        abs_path = os.path.join(FOLDER, filename)
        pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
        try:
            hwp.Open(abs_path, "HWP", "ForceOpen:1")
            hwp.SaveAs(pdf_path, "PDF", "")
            hwp.Clear(1)
            print(f"[OK] {filename}")
            success += 1
        except Exception as e:
            print(f"[FAIL] {filename}: {e}")
            fail += 1

    hwp.Quit()
    print(f"\n완료: 성공 {success}개, 실패 {fail}개")


if __name__ == "__main__":
    main()
