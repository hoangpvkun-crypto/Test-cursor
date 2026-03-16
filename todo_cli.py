#!/usr/bin/env python3
"""Ứng dụng To-Do List chạy trên CLI."""


def in_menu() -> str:
    print("\n=== QUẢN LÝ CÔNG VIỆC ===")
    print("1. Thêm công việc mới")
    print("2. Xóa công việc đã xong")
    print("3. In danh sách việc cần làm")
    print("0. Thoát")
    return input("Chọn chức năng: ").strip()


def them_cong_viec(ds_cong_viec: list[dict[str, str]]) -> None:
    ten_viec = input("Nhập nội dung công việc: ").strip()
    if not ten_viec:
        print("⚠️ Công việc không được để trống.")
        return

    ds_cong_viec.append({"ten": ten_viec})
    print("✅ Đã thêm công việc mới.")


def xoa_cong_viec(ds_cong_viec: list[dict[str, str]]) -> None:
    if not ds_cong_viec:
        print("⚠️ Danh sách đang trống, không có gì để xóa.")
        return

    in_danh_sach(ds_cong_viec)
    chi_so = input("Nhập số thứ tự công việc đã xong để xóa: ").strip()

    if not chi_so.isdigit():
        print("❌ Vui lòng nhập một số hợp lệ.")
        return

    vi_tri = int(chi_so) - 1
    if vi_tri < 0 or vi_tri >= len(ds_cong_viec):
        print("❌ Số thứ tự không tồn tại.")
        return

    da_xoa = ds_cong_viec.pop(vi_tri)
    print(f"✅ Đã xóa: {da_xoa['ten']}")


def in_danh_sach(ds_cong_viec: list[dict[str, str]]) -> None:
    print("\n--- DANH SÁCH CÔNG VIỆC ---")
    if not ds_cong_viec:
        print("(Trống)")
        return

    for i, cv in enumerate(ds_cong_viec, start=1):
        print(f"{i}. {cv['ten']}")


def chay_ung_dung() -> None:
    ds_cong_viec: list[dict[str, str]] = []

    while True:
        lua_chon = in_menu()

        if lua_chon == "1":
            them_cong_viec(ds_cong_viec)
        elif lua_chon == "2":
            xoa_cong_viec(ds_cong_viec)
        elif lua_chon == "3":
            in_danh_sach(ds_cong_viec)
        elif lua_chon == "0":
            print("Tạm biệt!")
            break
        else:
            print("❌ Lựa chọn không hợp lệ, vui lòng thử lại.")


if __name__ == "__main__":
    chay_ung_dung()
