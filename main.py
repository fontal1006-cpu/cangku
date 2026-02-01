import flet as ft
import openpyxl
import datetime
import os
import shutil
import time

def main(page: ft.Page):
    page.title = "🦅 小鹰提货明细生成器"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window_width = 380
    page.window_height = 700
    page.scroll = "auto"
    page.padding = 20

    # 配置路径（针对打包后的内部路径）
    DATA_PATH = "assets/data.xlsx"
    TPL_PATH = "assets/template.xlsx"
    CACHE_DIR = "temp_cache"

    if not os.path.exists(CACHE_DIR):
        os.makedirs(CACHE_DIR)

    # UI 变量
    search_input = ft.TextField(label="🔍 客户关键字", variant=ft.IndicatorCode.UNDERLINE, border_color="blue")
    product_input = ft.TextField(label="📦 产品名称", variant=ft.IndicatorCode.UNDERLINE)
    count_input = ft.TextField(label="📊 件数", variant=ft.IndicatorCode.UNDERLINE, keyboard_type=ft.KeyboardType.NUMBER)
    temp_dropdown = ft.SegmentedButton(
        segments=[
            ft.Segment(value="常温", label=ft.Text("常温")),
            ft.Segment(value="冷链", label=ft.Text("冷链")),
        ],
        selected={"常温"}
    )
    status_text = ft.Text("", color="gray")

    def clean_cache():
        """清理缓存文件夹"""
        for filename in os.listdir(CACHE_DIR):
            file_path = os.path.join(CACHE_DIR, filename)
            try:
                if os.path.isfile(file_path) or os.path.is_link(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"清理失败: {e}")

    def search_customer(keyword):
        if not os.path.exists(DATA_PATH):
            return None
        try:
            wb = openpyxl.load_workbook(DATA_PATH, data_only=True)
            ws = wb["Sheet2"]
            matches = {}
            for row in range(1, ws.max_row + 1):
                cell_value = ws.cell(row=row, column=2).value
                if cell_value and keyword in str(cell_value):
                    info = [
                        cell_value,
                        ws.cell(row=row + 1, column=2).value,
                        ws.cell(row=row + 2, column=2).value,
                        ws.cell(row=row + 3, column=2).value
                    ]
                    matches[str(cell_value)] = info
            wb.close()
            return matches
        except:
            return None

    def handle_generate(e):
        keyword = search_input.value.strip()
        if not keyword:
            page.snack_bar = ft.SnackBar(ft.Text("请输入关键字"))
            page.snack_bar.open = True
            page.update()
            return

        status_text.value = "🔍 正在检索客户..."
        page.update()

        matches = search_customer(keyword)
        if not matches:
            status_text.value = "❌ 未找到客户"
            page.update()
            return

        if len(matches) == 1:
            process_excel(list(matches.values())[0])
        else:
            # 多选列表
            def select_and_go(name):
                dlg.open = False
                process_excel(matches[name])

            list_items = [ft.ListTile(title=ft.Text(n), on_click=lambda _, n=n: select_and_go(n)) for n in matches.keys()]
            dlg = ft.AlertDialog(title=ft.Text("请选择精确客户"), content=ft.Column(list_items, tight=True))
            page.dialog = dlg
            dlg.open = True
            page.update()

    def process_excel(info):
        try:
            status_text.value = "📝 正在生成表格..."
            page.update()

            # 清理旧缓存
            clean_cache()

            # 打开模板
            wb = openpyxl.load_workbook(TPL_PATH)
            ws = wb["1"]

            # 填写数据
            today = datetime.datetime.now()
            ws["C2"] = today.strftime("%Y年%m月%d日")
            ws["B6"], ws["E6"], ws["C6"], ws["D6"] = info[0], info[1], info[2], info[3]
            ws["G6"] = product_input.value
            ws["J6"] = count_input.value
            ws["M6"] = list(temp_dropdown.selected)[0]

            # 保存到临时缓存
            filename = f"提货明细_{info[0]}_{today.strftime('%m%d%H%M')}.xlsx"
            temp_file_path = os.path.abspath(os.path.join(CACHE_DIR, filename))
            wb.save(temp_file_path)
            wb.close()

            status_text.value = "✅ 生成成功，准备分享"
            page.update()

            # 唤起手机分享
            page.share_files([temp_file_path])
            
            # 延时一点时间后清理（确保分享动作已读取文件）
            time.sleep(2)
            clean_cache()
            status_text.value = "🧹 缓存已安全清理"
            page.update()

        except Exception as ex:
            status_text.value = f"错误: {str(ex)}"
            page.update()

    # UI 布局
    page.add(
        ft.Column([
            ft.Container(
                content=ft.Text("🦅 小鹰提货生成器", size=28, weight="bold", color="blue"),
                alignment=ft.alignment.center,
                padding=20
            ),
            search_input,
            product_input,
            count_input,
            ft.Text("🌡️ 选择温度:"),
            temp_dropdown,
            ft.Divider(height=20, color="transparent"),
            ft.ElevatedButton(
                "🚀 生成并分享",
                on_click=handle_generate,
                style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=10)),
                width=400,
                height=50
            ),
            ft.Container(status_text, alignment=ft.alignment.center)
        ])
    )

ft.app(target=main, assets_dir="assets")
