#!/usr/bin/env python3

"""Ứng dụng To-Do List CLI với giao diện dạng app và chức năng mở rộng."""

from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path

DUONG_DAN_DU_LIEU = Path("todo_data.json")


def xoa_man_hinh() -> None:
    os.system("cls" if os.name == "nt" else "clear")


def ve_tieu_de() -> None:
    xoa_man_hinh()
    print("╔══════════════════════════════════════════════╗")
    print("║            📋 TO-DO LIST CLI APP             ║")
    print("╚══════════════════════════════════════════════╝")
    print(f"Thời gian: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")


def tai_du_lieu() -> list[dict[str, object]]:
    if not DUONG_DAN_DU_LIEU.exists():
        return []

    try:
        noi_dung = DUONG_DAN_DU_LIEU.read_text(encoding="utf-8")
        du_lieu = json.loads(noi_dung)
        if isinstance(du_lieu, list):
            return du_lieu
    except (json.JSONDecodeError, OSError):
        pass

    return []


def luu_du_lieu(ds_cong_viec: list[dict[str, object]]) -> None:
    DUONG_DAN_DU_LIEU.write_text(
        json.dumps(ds_cong_viec, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def tao_id_moi(ds_cong_viec: list[dict[str, object]]) -> int:
    if not ds_cong_viec:
        return 1
    return max(int(cv["id"]) for cv in ds_cong_viec) + 1


def tam_dung() -> None:
    input("\nNhấn Enter để tiếp tục...")


def in_menu() -> str:
    print("\n┌──────────── MENU CHỨC NĂNG ────────────┐")
    print("│ 1. Thêm công việc                      │")
    print("│ 2. Sửa công việc                       │")
    print("│ 3. Đánh dấu hoàn thành                 │")
    print("│ 4. Xóa công việc                       │")
    print("│ 5. Xem toàn bộ danh sách               │")
    print("│ 6. Tìm kiếm công việc                  │")
    print("│ 7. Lọc theo trạng thái                 │")
    print("│ 8. Thống kê nhanh                      │")
    print("│ 9. Xóa tất cả việc đã hoàn thành       │")
    print("│ 0. Thoát                               │")
    print("└────────────────────────────────────────┘")
    return input("👉 Chọn chức năng: ").strip()


def dinh_dang_trang_thai(cv: dict[str, object]) -> str:
    return "✅ Xong" if cv.get("hoan_thanh") else "⌛ Chưa xong"


def in_danh_sach(ds_cong_viec: list[dict[str, object]]) -> None:
    print("\n════════════════ DANH SÁCH CÔNG VIỆC ════════════════")
    if not ds_cong_viec:
        print("(Danh sách trống)")
        return

    print(f"{'STT':<5}{'ID':<5}{'Tên công việc':<30}{'Ưu tiên':<10}{'Trạng thái'}")
    print("-" * 70)
    for stt, cv in enumerate(ds_cong_viec, start=1):
        ten = str(cv.get("ten", ""))[:28]
        uu_tien = str(cv.get("uu_tien", "Vừa"))
        trang_thai = dinh_dang_trang_thai(cv)
        print(f"{stt:<5}{cv.get('id', ''):<5}{ten:<30}{uu_tien:<10}{trang_thai}")


def them_cong_viec(ds_cong_viec: list[dict[str, object]]) -> None:
    print("\n➕ THÊM CÔNG VIỆC MỚI")
    ten_viec = input("- Nhập nội dung: ").strip()

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
       main
    if not ten_viec:
        print("⚠️ Công việc không được để trống.")
        return


    uu_tien = input("- Ưu tiên (Thấp/Vừa/Cao, mặc định Vừa): ").strip().title()
    if uu_tien not in {"Thấp", "Vừa", "Cao"}:
        uu_tien = "Vừa"

    cong_viec = {
        "id": tao_id_moi(ds_cong_viec),
        "ten": ten_viec,
        "uu_tien": uu_tien,
        "hoan_thanh": False,
    }
    ds_cong_viec.append(cong_viec)
    luu_du_lieu(ds_cong_viec)
    print("✅ Đã thêm công việc.")


def chon_theo_stt(ds_cong_viec: list[dict[str, object]], thong_diep: str) -> int | None:
    if not ds_cong_viec:
        print("⚠️ Danh sách trống.")
        return None

    in_danh_sach(ds_cong_viec)
    du_lieu = input(thong_diep).strip()
    if not du_lieu.isdigit():
        print("❌ Vui lòng nhập số hợp lệ.")
        return None

    vi_tri = int(du_lieu) - 1
    if vi_tri < 0 or vi_tri >= len(ds_cong_viec):
        print("❌ STT không tồn tại.")
        return None

    return vi_tri


def sua_cong_viec(ds_cong_viec: list[dict[str, object]]) -> None:
    print("\n✏️ SỬA CÔNG VIỆC")
    vi_tri = chon_theo_stt(ds_cong_viec, "- Nhập STT cần sửa: ")
    if vi_tri is None:
        return

    cv = ds_cong_viec[vi_tri]
    ten_moi = input(f"- Tên mới (Enter để giữ '{cv['ten']}'): ").strip()
    uu_tien_moi = input(f"- Ưu tiên mới (Thấp/Vừa/Cao, hiện tại {cv['uu_tien']}): ").strip().title()

    if ten_moi:
        cv["ten"] = ten_moi
    if uu_tien_moi in {"Thấp", "Vừa", "Cao"}:
        cv["uu_tien"] = uu_tien_moi

    luu_du_lieu(ds_cong_viec)
    print("✅ Đã cập nhật công việc.")


def danh_dau_hoan_thanh(ds_cong_viec: list[dict[str, object]]) -> None:
    print("\n✅ ĐÁNH DẤU HOÀN THÀNH")
    vi_tri = chon_theo_stt(ds_cong_viec, "- Nhập STT đã hoàn thành: ")
    if vi_tri is None:
        return

    ds_cong_viec[vi_tri]["hoan_thanh"] = True
    luu_du_lieu(ds_cong_viec)
    print("🎉 Công việc đã được đánh dấu hoàn thành.")


def xoa_cong_viec(ds_cong_viec: list[dict[str, object]]) -> None:
    print("\n🗑️ XÓA CÔNG VIỆC")
    vi_tri = chon_theo_stt(ds_cong_viec, "- Nhập STT cần xóa: ")
    if vi_tri is None:
        return

    da_xoa = ds_cong_viec.pop(vi_tri)
    luu_du_lieu(ds_cong_viec)
    print(f"✅ Đã xóa: {da_xoa['ten']}")


def tim_kiem_cong_viec(ds_cong_viec: list[dict[str, object]]) -> None:
    print("\n🔎 TÌM KIẾM")
    tu_khoa = input("- Nhập từ khóa: ").strip().lower()
    if not tu_khoa:
        print("⚠️ Từ khóa không được để trống.")
        return

    ket_qua = [cv for cv in ds_cong_viec if tu_khoa in str(cv.get("ten", "")).lower()]
    in_danh_sach(ket_qua)


def loc_theo_trang_thai(ds_cong_viec: list[dict[str, object]]) -> None:
    print("\n📌 LỌC TRẠNG THÁI")
    print("1. Chưa hoàn thành")
    print("2. Đã hoàn thành")
    chon = input("- Chọn: ").strip()

    if chon == "1":
        ket_qua = [cv for cv in ds_cong_viec if not cv.get("hoan_thanh")]
    elif chon == "2":
        ket_qua = [cv for cv in ds_cong_viec if cv.get("hoan_thanh")]
    else:
        print("❌ Lựa chọn không hợp lệ.")
        return

    in_danh_sach(ket_qua)


def thong_ke(ds_cong_viec: list[dict[str, object]]) -> None:
    tong = len(ds_cong_viec)
    da_xong = sum(1 for cv in ds_cong_viec if cv.get("hoan_thanh"))
    chua_xong = tong - da_xong

    print("\n📊 THỐNG KÊ")
    print(f"- Tổng công việc   : {tong}")
    print(f"- Đã hoàn thành    : {da_xong}")
    print(f"- Chưa hoàn thành  : {chua_xong}")


def xoa_viec_da_xong(ds_cong_viec: list[dict[str, object]]) -> None:
    print("\n🧹 XÓA TOÀN BỘ VIỆC ĐÃ HOÀN THÀNH")
    so_luong_cu = len(ds_cong_viec)
    ds_cong_viec[:] = [cv for cv in ds_cong_viec if not cv.get("hoan_thanh")]
    da_xoa = so_luong_cu - len(ds_cong_viec)
    luu_du_lieu(ds_cong_viec)
    print(f"✅ Đã xóa {da_xoa} công việc đã hoàn thành.")


def chay_ung_dung() -> None:
    ds_cong_viec = tai_du_lieu()

    while True:
        ve_tieu_de()

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
        main
        lua_chon = in_menu()

        if lua_chon == "1":
            them_cong_viec(ds_cong_viec)
        elif lua_chon == "2":
            sua_cong_viec(ds_cong_viec)
        elif lua_chon == "3":
            danh_dau_hoan_thanh(ds_cong_viec)
        elif lua_chon == "4":
            xoa_cong_viec(ds_cong_viec)
        elif lua_chon == "5":
            in_danh_sach(ds_cong_viec)
        elif lua_chon == "6":
            tim_kiem_cong_viec(ds_cong_viec)
        elif lua_chon == "7":
            loc_theo_trang_thai(ds_cong_viec)
        elif lua_chon == "8":
            thong_ke(ds_cong_viec)
        elif lua_chon == "9":
            xoa_viec_da_xong(ds_cong_viec)
        elif lua_chon == "0":
            print("👋 Tạm biệt!")
            break
        else:
            print("❌ Lựa chọn không hợp lệ.")

        tam_dung()

            xoa_cong_viec(ds_cong_viec)
        elif lua_chon == "3":
            in_danh_sach(ds_cong_viec)
        elif lua_chon == "0":
            print("Tạm biệt!")
            break
        else:
            print("❌ Lựa chọn không hợp lệ, vui lòng thử lại.")
         main


if __name__ == "__main__":
    chay_ung_dung()
