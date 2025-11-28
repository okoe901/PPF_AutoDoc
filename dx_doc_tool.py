import os
import zipfile
import shutil
import subprocess
import tempfile
import json
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# =========================================================
#  ソリューション Zip 解凍 ＋ .msapp 展開
# =========================================================

def unzip_solution(zip_path: str, output_root: str):
    """
    PowerPlatform ソリューション Zip を解凍し、
    CanvasApps 内の .msapp を自動で zip 化して中身も展開する。
    出力先は output_root/solution/ 以下。
    """

    if not os.path.exists(zip_path):
        raise FileNotFoundError(f"ソリューションZipが見つかりません: {zip_path}")

    solution_dir = os.path.join(output_root, "solution")

    # 既存ディレクトリがあればそのまま上書き前提（必要なら削除ロジック追加）
    os.makedirs(solution_dir, exist_ok=True)

    # まずソリューションZipを解凍
    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(solution_dir)

    # CanvasApps フォルダを探索
    canvas_dir = os.path.join(solution_dir, "CanvasApps")
    if os.path.exists(canvas_dir):
        for file in os.listdir(canvas_dir):
            if file.endswith(".msapp"):
                msapp_path = os.path.join(canvas_dir, file)
                app_name = os.path.splitext(file)[0]
                app_extract_dir = os.path.join(canvas_dir, f"{app_name}_extracted")

                os.makedirs(app_extract_dir, exist_ok=True)

                # .msapp を一時的に .zip として扱う
                temp_zip = msapp_path + ".zip"
                shutil.copy(msapp_path, temp_zip)

                try:
                    with zipfile.ZipFile(temp_zip, "r") as z:
                        z.extractall(app_extract_dir)
                except zipfile.BadZipFile:
                    print(f"[WARN] {file} は Zip として展開できませんでした。")
                finally:
                    if os.path.exists(temp_zip):
                        os.remove(temp_zip)

    return solution_dir


# =========================================================
#  SharePoint リスト設計書（PnP.PowerShell 経由）
# =========================================================

def get_sp_list_columns(site_url: str, list_name: str):
    """
    PowerShell（PnP.PowerShell）を実行し、
    SharePoint リストの列構造を JSON で取得する。

    ※ 前提：
        PowerShell で以下インストール済み:
        Install-Module -Name PnP.PowerShell -Scope CurrentUser
    """

    temp_json = os.path.join(tempfile.gettempdir(), "sp_list_fields.json")

    # PowerShell スクリプトを組み立て
    # -Interactive でブラウザ認証
    ps_script = f"""
    Import-Module PnP.PowerShell
    Connect-PnPOnline -Url "{site_url}" -Interactive
    $fields = Get-PnPField -List "{list_name}"
    $fields | ConvertTo-Json -Depth 10 | Out-File "{temp_json}" -Encoding UTF8
    """

    # PowerShell 実行
    # Windows 前提。PowerShell 7 なら "pwsh" に変更。
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            capture_output=True,
            text=True,
            check=False
        )
    except FileNotFoundError:
        raise RuntimeError("PowerShell が見つかりません。Windows で実行しているか確認してください。")

    if result.returncode != 0:
        # エラー内容を投げる
        raise RuntimeError(
            f"PowerShell 実行中にエラーが発生しました。\n"
            f"STDOUT: {result.stdout}\n"
            f"STDERR: {result.stderr}"
        )

    if not os.path.exists(temp_json):
        raise RuntimeError("フィールド情報のJSONが出力されませんでした。権限やリスト名を確認してください。")

    with open(temp_json, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Get-PnPField の戻りが単一オブジェクトのこともあるのでリストに正規化
    if isinstance(data, dict):
        data = [data]

    return data


def export_sp_list_to_csv(fields, output_root: str, list_name: str):
    """
    SharePoint のフィールド情報(JSONオブジェクト)から
    設計書用の CSV を出力する。
    出力先: output_root/sharepoint/{list_name}_fields.csv
    """

    sp_dir = os.path.join(output_root, "sharepoint")
    os.makedirs(sp_dir, exist_ok=True)

    safe_list_name = list_name.replace("/", "_").replace("\\", "_")
    csv_path = os.path.join(sp_dir, f"{safe_list_name}_fields.csv")
    json_path = os.path.join(sp_dir, f"{safe_list_name}_fields.json")

    # JSON も保存しておく（後で GPT に投げる用）
    with open(json_path, "w", encoding="utf-8") as jf:
        json.dump(fields, jf, ensure_ascii=False, indent=2)

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["表示名", "内部名", "型", "必須", "説明", "DefaultValue", "その他"])

        for col in fields:
            title = col.get("Title") or col.get("DisplayName") or ""
            internal_name = col.get("InternalName") or ""
            type_str = col.get("TypeAsString") or col.get("TypeDisplayName") or ""
            required = col.get("Required")
            desc = col.get("Description") or ""
            default = col.get("DefaultValue") or ""

            # ここにOption/Lookupなどの追加情報をまとめて入れる
            extra = []

            # 選択肢（Choice）系
            if "Choices" in col and col["Choices"]:
                extra.append(f"Choices: {', '.join(col['Choices'])}")

            # Lookup系
            lookup_list = col.get("LookupList")
            lookup_field = col.get("LookupField")
            if lookup_list or lookup_field:
                extra.append(f"Lookup: List={lookup_list}, Field={lookup_field}")

            extra_str = " | ".join(extra)

            writer.writerow([
                title,
                internal_name,
                type_str,
                required,
                desc,
                default,
                extra_str
            ])

    return csv_path, json_path


# =========================================================
#  GPT 用プロンプト生成
# =========================================================

def generate_gpt_prompt(output_root: str):
    """
    ソリューション・SharePoint の出力フォルダ構成を前提に、
    GPT にそのままコピペできるプロンプトを生成する。
    出力先: output_root/gpt_prompt.txt
    """

    solution_dir = os.path.join(output_root, "solution")
    sp_dir = os.path.join(output_root, "sharepoint")

    # フォルダがあるか軽く確認（なければ注意文だけ入れる）
    has_solution = os.path.exists(solution_dir)
    has_sp = os.path.exists(sp_dir)

    prompt_lines = []

    prompt_lines.append("あなたはプロのシステム設計書エンジニアです。")
    prompt_lines.append("以下のファイルやフォルダに含まれる情報を使って、PowerPlatform ＋ SharePoint の設計書を作成してください。")
    prompt_lines.append("")
    prompt_lines.append("▼ 分析対象となるファイル・フォルダの場所（ローカルパス）")
    if has_solution:
        prompt_lines.append(f"- PowerPlatform ソリューション展開フォルダ: {solution_dir}")
        prompt_lines.append("  - CanvasApps 配下の *_extracted フォルダ内 Controls.json / Properties.json")
        prompt_lines.append("  - Workflows フォルダ内の Flow 定義 (JSON)")
        prompt_lines.append("  - customizations.xml に含まれる Dataverse テーブル定義")
    else:
        prompt_lines.append("- (solution フォルダがまだ生成されていません)")

    if has_sp:
        prompt_lines.append(f"- SharePoint リスト設計書フォルダ: {sp_dir}")
        prompt_lines.append("  - *_fields.csv : 各リストの列設計書")
        prompt_lines.append("  - *_fields.json : 生のフィールド情報（必要に応じて参照）")
    else:
        prompt_lines.append("- (sharepoint フォルダがまだ生成されていません)")

    prompt_lines.append("")
    prompt_lines.append("▼ 出力してほしいドキュメント")
    prompt_lines.append("【1】Canvas Apps 画面設計書")
    prompt_lines.append("- 画面一覧（画面名・用途）")
    prompt_lines.append("- 各画面に配置されている主要コントロール一覧（Button, Gallery など）")
    prompt_lines.append("- 各コントロールの主なイベント（OnSelect, Items, Visible など）とその役割")
    prompt_lines.append("- グローバル変数(Set)・コンテキスト変数(UpdateContext)・コレクション(ClearCollect)の一覧と用途")
    prompt_lines.append("")
    prompt_lines.append("【2】Power Automate フロー設計書")
    prompt_lines.append("- 各フローごとに、トリガー・アクション・条件分岐・ループ処理の概要")
    prompt_lines.append("- どの Canvas App から呼ばれているか、その関係性")
    prompt_lines.append("")
    prompt_lines.append("【3】Dataverse テーブル設計書")
    prompt_lines.append("- テーブル一覧")
    prompt_lines.append("- 各テーブルの列（表示名・スキーマ名・型・必須・説明）")
    prompt_lines.append("- Lookup 関係、OptionSet の選択肢などの関係情報")
    prompt_lines.append("")
    prompt_lines.append("【4】SharePoint リスト設計書")
    prompt_lines.append("- リストごとに、列（表示名・内部名・型・必須・説明・選択肢・Lookup）の一覧")
    prompt_lines.append("- Dataverse / Canvas App / Flow と連携している場合、その関係性")
    prompt_lines.append("")
    prompt_lines.append("【5】全体構成図（テキストベースでOK）")
    prompt_lines.append("- アプリ、フロー、Dataverse テーブル、SharePoint リストの関連図（どこからどこへデータが流れているか）")
    prompt_lines.append("")
    prompt_lines.append("出力形式は Markdown で、見出しや表を使って読みやすく整理してください。")

    prompt_text = "\n".join(prompt_lines)

    os.makedirs(output_root, exist_ok=True)
    prompt_path = os.path.join(output_root, "gpt_prompt.txt")
    with open(prompt_path, "w", encoding="utf-8") as f:
        f.write(prompt_text)

    return prompt_path


# =========================================================
#  Tkinter UI
# =========================================================

class DXDocToolApp:
    def __init__(self, master):
        self.master = master
        master.title("DX ドキュメント自動化ツール（PowerPlatform ＋ SharePoint）")
        master.geometry("720x480")
        master.resizable(False, False)

        self.solution_zip_var = tk.StringVar()
        self.output_root_var = tk.StringVar()
        self.sp_site_url_var = tk.StringVar()
        self.sp_list_name_var = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        notebook = ttk.Notebook(self.master)
        notebook.pack(fill="both", expand=True)

        # --- タブ1: ソリューションZip ---
        frame_sol = ttk.Frame(notebook, padding=15)
        notebook.add(frame_sol, text="1. ソリューション解析")

        ttk.Label(frame_sol, text="ソリューションZipファイル").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_sol, textvariable=self.solution_zip_var, width=60).grid(row=1, column=0, sticky="w")
        ttk.Button(frame_sol, text="選択", command=self.select_solution_zip)\
            .grid(row=1, column=1, padx=5)

        ttk.Label(frame_sol, text="出力ルートフォルダ").grid(row=2, column=0, sticky="w", pady=(15, 0))
        ttk.Entry(frame_sol, textvariable=self.output_root_var, width=60).grid(row=3, column=0, sticky="w")
        ttk.Button(frame_sol, text="選択", command=self.select_output_root)\
            .grid(row=3, column=1, padx=5)

        ttk.Button(frame_sol, text="ソリューションZip解凍 ＋ .msapp 展開を実行",
                   command=self.run_solution_unzip)\
            .grid(row=4, column=0, columnspan=2, pady=30)

        ttk.Label(frame_sol, text="※ CanvasApps 内の .msapp も自動で展開されます。").grid(row=5, column=0, columnspan=2, sticky="w")

        # --- タブ2: SharePoint リスト設計書 ---
        frame_sp = ttk.Frame(notebook, padding=15)
        notebook.add(frame_sp, text="2. SharePoint 設計書")

        ttk.Label(frame_sp, text="SharePoint サイト URL").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_sp, textvariable=self.sp_site_url_var, width=60).grid(row=1, column=0, sticky="w", columnspan=2)

        ttk.Label(frame_sp, text="リスト名（表示名）").grid(row=2, column=0, sticky="w", pady=(15, 0))
        ttk.Entry(frame_sp, textvariable=self.sp_list_name_var, width=40).grid(row=3, column=0, sticky="w")

        ttk.Button(frame_sp, text="このリストの設計書を取得（ブラウザで認証）",
                   command=self.run_sharepoint_extract)\
            .grid(row=4, column=0, columnspan=2, pady=25)

        ttk.Label(frame_sp, text="※ PnP.PowerShell を利用します（事前に Install-Module が必要）。")\
            .grid(row=5, column=0, columnspan=2, sticky="w")

        # --- タブ3: GPT プロンプト生成 ---
        frame_gpt = ttk.Frame(notebook, padding=15)
        notebook.add(frame_gpt, text="3. GPT プロンプト")

        ttk.Label(frame_gpt, text="出力ルートフォルダ").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_gpt, textvariable=self.output_root_var, width=60).grid(row=1, column=0, sticky="w")
        ttk.Button(frame_gpt, text="選択", command=self.select_output_root)\
            .grid(row=1, column=1, padx=5)

        ttk.Button(frame_gpt, text="GPT 用プロンプトファイルを生成（gpt_prompt.txt）",
                   command=self.run_generate_prompt)\
            .grid(row=2, column=0, columnspan=2, pady=30)

        ttk.Label(frame_gpt, text="※ 生成された gpt_prompt.txt を GPT にコピペすると、設計書を自動生成しやすくなります。")\
            .grid(row=3, column=0, columnspan=2, sticky="w")

    # ------------- イベントハンドラ -------------

    def select_solution_zip(self):
        path = filedialog.askopenfilename(
            title="ソリューションZipを選択",
            filetypes=[("Solution Zip", "*.zip"), ("All files", "*.*")]
        )
        if path:
            self.solution_zip_var.set(path)

    def select_output_root(self):
        path = filedialog.askdirectory(title="出力ルートフォルダを選択")
        if path:
            self.output_root_var.set(path)

    def run_solution_unzip(self):
        zip_path = self.solution_zip_var.get()
        output_root = self.output_root_var.get()

        if not zip_path:
            messagebox.showerror("エラー", "ソリューションZipファイルを選択してください。")
            return

        if not output_root:
            messagebox.showerror("エラー", "出力ルートフォルダを選択してください。")
            return

        try:
            solution_dir = unzip_solution(zip_path, output_root)
            messagebox.showinfo("完了", f"ソリューションZipの展開が完了しました。\n\n出力先: {solution_dir}")
        except Exception as e:
            messagebox.showerror("エラー", f"ソリューションの展開中にエラーが発生しました。\n\n{e}")

    def run_sharepoint_extract(self):
        site_url = self.sp_site_url_var.get().strip()
        list_name = self.sp_list_name_var.get().strip()
        output_root = self.output_root_var.get().strip()

        if not output_root:
            messagebox.showerror("エラー", "出力ルートフォルダを先に指定してください。（タブ1またはタブ3）")
            return

        if not site_url:
            messagebox.showerror("エラー", "SharePoint サイト URL を入力してください。")
            return

        if not list_name:
            messagebox.showerror("エラー", "リスト名を入力してください。")
            return

        try:
            fields = get_sp_list_columns(site_url, list_name)
            csv_path, json_path = export_sp_list_to_csv(fields, output_root, list_name)
            messagebox.showinfo(
                "完了",
                f"SharePoint リスト設計書の取得が完了しました。\n\nCSV: {csv_path}\nJSON: {json_path}\n\nGPT での設計書生成時に参照してください。"
            )
        except Exception as e:
            messagebox.showerror("エラー", f"SharePoint リスト取得中にエラーが発生しました。\n\n{e}")

    def run_generate_prompt(self):
        output_root = self.output_root_var.get().strip()

        if not output_root:
            messagebox.showerror("エラー", "出力ルートフォルダを選択してください。")
            return

        try:
            prompt_path = generate_gpt_prompt(output_root)
            messagebox.showinfo(
                "完了",
                f"GPT 用プロンプトファイルを生成しました。\n\n{prompt_path}\n\nこのファイルの中身を GPT にコピペして使ってください。"
            )
        except Exception as e:
            messagebox.showerror("エラー", f"プロンプト生成中にエラーが発生しました。\n\n{e}")


# =========================================================
#  エントリポイント
# =========================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = DXDocToolApp(root)
    root.mainloop()
