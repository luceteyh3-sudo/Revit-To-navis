"""
Revit → Navisworks 일괄 변환기
- Revit(.rvt) 파일을 Navisworks(.nwc) 파일로 일괄 변환
- 원본 Revit 파일은 변환하지 않고 NWC 파일만 별도 경로에 저장
- 하위 폴더 구조를 유지하며 저장
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import subprocess
import threading
from pathlib import Path
import sys
import json

SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".revit_navis_converter.json")


class RevitToNavisApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Revit → Navisworks 일괄 변환기")
        self.root.geometry("900x680")
        self.root.minsize(720, 520)

        # 상태 변수
        self.navisworks_path = tk.StringVar()
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.is_converting = False
        self.stop_flag = False
        self.revit_files: list[str] = []
        self.total_files = 0
        self.converted_count = 0
        self.failed_count = 0

        self._build_ui()
        self._auto_detect_navisworks()
        self._load_settings()

    # ──────────────────────────────────────────────
    # UI 구성
    # ──────────────────────────────────────────────

    def _build_ui(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Success.TLabel", foreground="green")
        style.configure("Error.TLabel", foreground="red")
        style.configure("Info.TLabel", foreground="blue")

        main = ttk.Frame(self.root, padding=12)
        main.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)

        # 제목
        ttk.Label(
            main,
            text="Revit → Navisworks 일괄 변환기",
            font=("Malgun Gothic", 15, "bold"),
        ).grid(row=0, column=0, columnspan=3, pady=(0, 12))

        # ── Navisworks 경로 ──
        ttk.Label(main, text="Navisworks 설치 경로:").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(main, textvariable=self.navisworks_path).grid(
            row=1, column=1, sticky="ew", padx=5
        )
        ttk.Button(main, text="찾아보기", command=self._browse_navisworks, width=10).grid(
            row=1, column=2, padx=3
        )

        # ── Revit 입력 경로 ──
        ttk.Label(main, text="Revit 파일 최상위 경로:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(main, textvariable=self.input_path).grid(
            row=2, column=1, sticky="ew", padx=5
        )
        ttk.Button(main, text="찾아보기", command=self._browse_input, width=10).grid(
            row=2, column=2, padx=3
        )

        # ── 저장 경로 ──
        ttk.Label(main, text="NWC 저장 경로:").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(main, textvariable=self.output_path).grid(
            row=3, column=1, sticky="ew", padx=5
        )
        ttk.Button(main, text="찾아보기", command=self._browse_output, width=10).grid(
            row=3, column=2, padx=3
        )

        # ── 파일 검색 버튼 ──
        search_row = ttk.Frame(main)
        search_row.grid(row=4, column=0, columnspan=3, pady=8, sticky="w")
        ttk.Button(
            search_row, text="  Revit 파일 검색  ", command=self._scan_files
        ).pack(side="left")
        self.file_count_label = ttk.Label(search_row, text="", foreground="gray")
        self.file_count_label.pack(side="left", padx=12)

        # ── 파일 목록 ──
        list_frame = ttk.LabelFrame(main, text="검색된 Revit 파일 목록", padding=5)
        list_frame.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=4)
        main.rowconfigure(5, weight=1)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        cols = ("파일명", "상대 경로", "상태")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", height=12)
        self.tree.heading("파일명", text="파일명")
        self.tree.heading("상대 경로", text="상대 경로")
        self.tree.heading("상태", text="상태")
        self.tree.column("파일명", width=220, minwidth=120)
        self.tree.column("상대 경로", width=480, minwidth=200)
        self.tree.column("상태", width=120, minwidth=80)

        sb_y = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        sb_x = ttk.Scrollbar(list_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb_y.grid(row=0, column=1, sticky="ns")
        sb_x.grid(row=1, column=0, sticky="ew")

        self.tree.tag_configure("success", foreground="green")
        self.tree.tag_configure("fail", foreground="red")
        self.tree.tag_configure("running", foreground="blue")

        # ── 진행률 ──
        self.progress_var = tk.DoubleVar()
        ttk.Progressbar(main, variable=self.progress_var, maximum=100).grid(
            row=6, column=0, columnspan=3, sticky="ew", pady=(6, 2)
        )

        self.status_label = ttk.Label(main, text="대기 중...", foreground="gray")
        self.status_label.grid(row=7, column=0, columnspan=3, sticky="w")

        # ── 버튼 행 ──
        btn_row = ttk.Frame(main)
        btn_row.grid(row=8, column=0, columnspan=3, pady=(10, 0))

        self.convert_btn = ttk.Button(
            btn_row, text="  변환 시작  ", command=self._start_conversion, width=14
        )
        self.convert_btn.pack(side="left", padx=5)

        self.stop_btn = ttk.Button(
            btn_row, text="중지", command=self._stop_conversion, width=8, state="disabled"
        )
        self.stop_btn.pack(side="left", padx=5)

        ttk.Button(btn_row, text="로그 저장", command=self._save_log, width=10).pack(
            side="left", padx=5
        )
        ttk.Button(btn_row, text="닫기", command=self._on_close, width=8).pack(
            side="right", padx=5
        )

    # ──────────────────────────────────────────────
    # Navisworks 자동 감지
    # ──────────────────────────────────────────────

    def _auto_detect_navisworks(self):
        """Navisworks 설치 경로 자동 탐색 (2019~2026)"""
        for year in range(2026, 2018, -1):
            for product in ["Manage", "Simulate"]:
                path = rf"C:\Program Files\Autodesk\Navisworks {product} {year}"
                filetool = os.path.join(path, "FiletoolsTaskRunner.exe")
                roamer = os.path.join(path, "roamer.exe")
                if os.path.exists(filetool) or os.path.exists(roamer):
                    self.navisworks_path.set(path)
                    return

    # ──────────────────────────────────────────────
    # 설정 저장/불러오기
    # ──────────────────────────────────────────────

    def _load_settings(self):
        try:
            if os.path.exists(SETTINGS_FILE):
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    s = json.load(f)
                if not self.navisworks_path.get() and s.get("navisworks_path"):
                    self.navisworks_path.set(s["navisworks_path"])
                if s.get("output_path"):
                    self.output_path.set(s["output_path"])
        except Exception:
            pass

    def _save_settings(self):
        try:
            s = {
                "navisworks_path": self.navisworks_path.get(),
                "output_path": self.output_path.get(),
            }
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(s, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ──────────────────────────────────────────────
    # 경로 찾아보기
    # ──────────────────────────────────────────────

    def _browse_navisworks(self):
        path = filedialog.askdirectory(title="Navisworks 설치 폴더 선택")
        if path:
            self.navisworks_path.set(path.replace("/", "\\"))

    def _browse_input(self):
        path = filedialog.askdirectory(title="Revit 파일이 있는 최상위 폴더 선택")
        if path:
            self.input_path.set(path.replace("/", "\\"))

    def _browse_output(self):
        path = filedialog.askdirectory(title="NWC 파일을 저장할 폴더 선택")
        if path:
            self.output_path.set(path.replace("/", "\\"))

    # ──────────────────────────────────────────────
    # 파일 검색
    # ──────────────────────────────────────────────

    def _scan_files(self):
        input_dir = self.input_path.get().strip()
        if not input_dir:
            messagebox.showwarning("경고", "Revit 파일 경로를 먼저 선택해주세요.")
            return
        if not os.path.isdir(input_dir):
            messagebox.showerror("오류", "선택한 경로가 존재하지 않습니다.")
            return

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.revit_files = []

        self._set_status("파일 검색 중...", "blue")
        self.root.update()

        for rvt_path in sorted(Path(input_dir).rglob("*.rvt")):
            self.revit_files.append(str(rvt_path))
            rel_path = str(rvt_path.relative_to(input_dir))
            self.tree.insert("", "end", values=(rvt_path.name, rel_path, "대기"))

        count = len(self.revit_files)
        self.total_files = count
        self.file_count_label.config(text=f"총 {count}개 파일 발견")
        self._set_status(f"{count}개의 Revit 파일을 찾았습니다.", "black")

    # ──────────────────────────────────────────────
    # 변환 시작
    # ──────────────────────────────────────────────

    def _start_conversion(self):
        if not self.revit_files:
            messagebox.showwarning("경고", "먼저 [Revit 파일 검색]을 실행해주세요.")
            return

        navis_path = self.navisworks_path.get().strip()
        output_dir = self.output_path.get().strip()

        if not navis_path:
            messagebox.showerror("오류", "Navisworks 설치 경로를 설정해주세요.")
            return
        if not output_dir:
            messagebox.showerror("오류", "NWC 저장 경로를 설정해주세요.")
            return

        filetool = os.path.join(navis_path, "FiletoolsTaskRunner.exe")
        roamer = os.path.join(navis_path, "roamer.exe")
        if not os.path.exists(filetool) and not os.path.exists(roamer):
            messagebox.showerror(
                "오류",
                f"Navisworks 변환 도구를 찾을 수 없습니다.\n\n"
                f"확인한 경로:\n  {filetool}\n  {roamer}\n\n"
                "Navisworks Manage 또는 Simulate가 올바르게 설치되어 있는지 확인해주세요.",
            )
            return

        self._save_settings()

        # 상태 초기화
        self.is_converting = True
        self.stop_flag = False
        self.converted_count = 0
        self.failed_count = 0
        self.progress_var.set(0)
        self.convert_btn.config(state="disabled")
        self.stop_btn.config(state="normal")

        for item in self.tree.get_children():
            v = self.tree.item(item)["values"]
            self.tree.item(item, values=(v[0], v[1], "대기"), tags=())

        thread = threading.Thread(
            target=self._conversion_worker,
            args=(navis_path, output_dir),
            daemon=True,
        )
        thread.start()

    # ──────────────────────────────────────────────
    # 변환 작업 (백그라운드 스레드)
    # ──────────────────────────────────────────────

    def _conversion_worker(self, navis_path: str, output_dir: str):
        input_dir = self.input_path.get().strip()
        items = self.tree.get_children()

        for idx, (rvt_path, item) in enumerate(zip(self.revit_files, items)):
            if self.stop_flag:
                self.root.after(0, self._set_status, "변환이 중지되었습니다.", "orange")
                break

            filename = os.path.basename(rvt_path)
            self.root.after(0, self._update_item_status, item, "변환 중...", "running")
            self.root.after(
                0,
                self._set_status,
                f"변환 중: {filename}  ({idx + 1}/{self.total_files})",
                "blue",
            )

            # 출력 경로 계산 (하위 폴더 구조 유지)
            try:
                rel_path = os.path.relpath(rvt_path, input_dir)
            except ValueError:
                rel_path = filename
            out_file = os.path.join(output_dir, os.path.splitext(rel_path)[0] + ".nwc")
            os.makedirs(os.path.dirname(out_file), exist_ok=True)

            success, message = self._convert_single_file(rvt_path, out_file, navis_path)

            if success:
                self.converted_count += 1
                self.root.after(0, self._update_item_status, item, "성공", "success")
            else:
                self.failed_count += 1
                short_msg = message[:40] if len(message) > 40 else message
                self.root.after(0, self._update_item_status, item, f"실패: {short_msg}", "fail")

            progress = (idx + 1) / self.total_files * 100
            self.root.after(0, self.progress_var.set, progress)

        self.root.after(0, self._conversion_complete)

    def _convert_single_file(
        self, rvt_path: str, out_file: str, navis_path: str
    ) -> tuple[bool, str]:
        """단일 .rvt → .nwc 변환"""
        filetool = os.path.join(navis_path, "FiletoolsTaskRunner.exe")
        roamer = os.path.join(navis_path, "roamer.exe")
        out_dir = os.path.dirname(out_file)

        if os.path.exists(filetool):
            # FiletoolsTaskRunner 방식 (Navisworks Manage 권장)
            cmd = [filetool, "-in", rvt_path, "-of", "nwc", "-out", out_file]
        elif os.path.exists(roamer):
            # roamer.exe 방식
            cmd = [roamer, "/i", rvt_path, "/of", "nwc", "/o", out_dir]
        else:
            return False, "변환 도구를 찾을 수 없습니다."

        no_window = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=600,  # 파일당 최대 10분
                creationflags=no_window,
            )
            # 파일 생성 여부로 성공 판단
            if os.path.exists(out_file):
                return True, "성공"
            if result.returncode == 0:
                return True, "성공 (출력 파일 미확인)"
            err = (result.stderr or result.stdout or "알 수 없는 오류").strip()
            return False, err
        except subprocess.TimeoutExpired:
            return False, "시간 초과 (10분 초과)"
        except Exception as exc:
            return False, str(exc)

    # ──────────────────────────────────────────────
    # UI 업데이트 헬퍼 (메인 스레드에서 호출)
    # ──────────────────────────────────────────────

    def _update_item_status(self, item, status: str, tag: str = ""):
        v = self.tree.item(item)["values"]
        self.tree.item(item, values=(v[0], v[1], status), tags=(tag,) if tag else ())

    def _set_status(self, text: str, color: str = "black"):
        self.status_label.config(text=text, foreground=color)

    def _conversion_complete(self):
        self.is_converting = False
        self.convert_btn.config(state="normal")
        self.stop_btn.config(state="disabled")

        summary = (
            f"변환 완료 — 성공: {self.converted_count}개 / "
            f"실패: {self.failed_count}개 / 전체: {self.total_files}개"
        )
        color = "green" if self.failed_count == 0 else "orange"
        self._set_status(summary, color)

        messagebox.showinfo(
            "변환 완료",
            f"변환이 완료되었습니다.\n\n"
            f"  성공: {self.converted_count}개\n"
            f"  실패: {self.failed_count}개\n"
            f"  전체: {self.total_files}개",
        )

    def _stop_conversion(self):
        self.stop_flag = True
        self.stop_btn.config(state="disabled")
        self._set_status("중지 요청 중... 현재 파일 완료 후 중지됩니다.", "orange")

    # ──────────────────────────────────────────────
    # 로그 저장
    # ──────────────────────────────────────────────

    def _save_log(self):
        log_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("텍스트 파일", "*.txt"), ("모든 파일", "*.*")],
            title="로그 파일 저장",
            initialfile="navis_convert_log.txt",
        )
        if not log_path:
            return
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                f.write("Revit → Navisworks 변환 로그\n")
                f.write("=" * 60 + "\n\n")
                f.write(f"Revit 경로 : {self.input_path.get()}\n")
                f.write(f"저장 경로  : {self.output_path.get()}\n\n")
                f.write(f"{'파일명':<40} {'상태'}\n")
                f.write("-" * 60 + "\n")
                for item in self.tree.get_children():
                    v = self.tree.item(item)["values"]
                    f.write(f"{v[1]:<40} {v[2]}\n")
                f.write(f"\n성공: {self.converted_count} / 실패: {self.failed_count} / 전체: {self.total_files}\n")
            messagebox.showinfo("저장 완료", f"로그가 저장되었습니다:\n{log_path}")
        except Exception as exc:
            messagebox.showerror("오류", f"로그 저장 실패:\n{exc}")

    def _on_close(self):
        if self.is_converting:
            if not messagebox.askyesno("확인", "변환이 진행 중입니다. 종료하시겠습니까?"):
                return
        self._save_settings()
        self.root.destroy()


# ──────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────

def main():
    root = tk.Tk()
    app = RevitToNavisApp(root)
    root.protocol("WM_DELETE_WINDOW", app._on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
