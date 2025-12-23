# tools/analyze_pdf_format.py
"""
分析 PDF 表格格式，提取儲存格反灰資訊
"""
import pdfplumber
from pathlib import Path
import sys

def analyze_pdf_tables(pdf_path: str, page_range: tuple = None):
    """分析 PDF 表格結構和格式"""
    with pdfplumber.open(pdf_path) as pdf:
        start = page_range[0] if page_range else 0
        end = page_range[1] if page_range else len(pdf.pages)

        for page_idx in range(start, min(end, len(pdf.pages))):
            page = pdf.pages[page_idx]
            page_num = page_idx + 1

            print(f"\n{'='*60}")
            print(f"Page {page_num}")
            print(f"{'='*60}")

            # 抽取表格
            tables = page.extract_tables({
                'vertical_strategy': 'lines',
                'horizontal_strategy': 'lines',
            })

            if not tables:
                print("  (無表格)")
                continue

            for t_idx, table in enumerate(tables):
                if not table:
                    continue

                print(f"\n--- Table {t_idx + 1} ---")
                print(f"Rows: {len(table)}, Cols: {max(len(r) for r in table if r)}")

                for r_idx, row in enumerate(table[:5]):  # 只顯示前 5 行
                    if row:
                        # 顯示每個 cell 的內容（截斷）
                        cells = []
                        for c in row:
                            cell_text = (str(c) if c else "").replace('\n', ' ')[:30]
                            cells.append(f"[{cell_text}]")
                        print(f"  Row {r_idx}: {' | '.join(cells)}")

                if len(table) > 5:
                    print(f"  ... (還有 {len(table) - 5} 行)")

            # 分析頁面上的矩形（可能是反灰背景）
            rects = page.rects
            if rects:
                print(f"\n--- Rectangles (可能是反灰背景) ---")
                print(f"  共 {len(rects)} 個矩形")

                # 檢查是否有填充色
                for rect in rects[:10]:
                    fill = rect.get('fill') or rect.get('non_stroking_color')
                    if fill:
                        print(f"  Fill color: {fill} at ({rect['x0']:.1f}, {rect['y0']:.1f})")


def analyze_page_chars_and_rects(pdf_path: str, page_num: int):
    """詳細分析特定頁面的字元和矩形"""
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_num - 1]

        print(f"\n{'='*60}")
        print(f"Page {page_num} 詳細分析")
        print(f"{'='*60}")

        # 抽取所有矩形
        rects = page.rects
        print(f"\n總共 {len(rects)} 個矩形")

        # 找出有填充色的矩形（反灰背景）
        filled_rects = []
        for rect in rects:
            fill = rect.get('fill') or rect.get('non_stroking_color')
            stroke = rect.get('stroking_color')
            if fill and fill != (1, 1, 1):  # 排除白色
                filled_rects.append({
                    'x0': rect['x0'],
                    'y0': rect['y0'],
                    'x1': rect['x1'],
                    'y1': rect['y1'],
                    'fill': fill,
                })

        print(f"有填充色（非白色）的矩形: {len(filled_rects)}")
        for i, r in enumerate(filled_rects[:20]):
            print(f"  [{i}] ({r['x0']:.1f}, {r['y0']:.1f}) - ({r['x1']:.1f}, {r['y1']:.1f}): fill={r['fill']}")

        # 分析表格結構
        tables = page.find_tables({
            'vertical_strategy': 'lines',
            'horizontal_strategy': 'lines',
        })

        print(f"\n找到 {len(tables)} 個表格")
        for t_idx, table in enumerate(tables):
            bbox = table.bbox
            print(f"\nTable {t_idx + 1} bbox: ({bbox[0]:.1f}, {bbox[1]:.1f}) - ({bbox[2]:.1f}, {bbox[3]:.1f})")

            # 顯示表格 cells
            cells = table.cells
            print(f"  Cells: {len(cells)}")

            # 計算哪些 cell 和填充矩形重疊
            for c_idx, cell in enumerate(cells[:10]):
                cx0, cy0, cx1, cy1 = cell
                # 檢查是否有填充矩形覆蓋此 cell
                for filled in filled_rects:
                    # 檢查重疊
                    if (filled['x0'] <= cx0 + 5 and filled['x1'] >= cx1 - 5 and
                        filled['y0'] <= cy0 + 5 and filled['y1'] >= cy1 - 5):
                        print(f"  Cell {c_idx} ({cx0:.1f}, {cy0:.1f})-({cx1:.1f}, {cy1:.1f}) 有背景色: {filled['fill']}")
                        break


if __name__ == "__main__":
    # 預設分析 DYS830.pdf
    pdf_path = Path(__file__).parent.parent / "templates" / "Samples" / "DYS830.pdf"

    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]

    print(f"分析 PDF: {pdf_path}")

    # 分析 page 5-7 (通常是能源概覽開始的地方)
    analyze_pdf_tables(str(pdf_path), (4, 10))

    # 詳細分析 page 5
    print("\n" + "="*60)
    print("詳細分析 Page 5")
    analyze_page_chars_and_rects(str(pdf_path), 5)
