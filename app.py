import easyocr
import pandas as pd
from pathlib import Path
import argparse
from copy import copy
from statistics import median
from PIL import Image


def image_to_excel(image_path: str, output_excel: str) -> None:
    image_file = Path(image_path)
    if not image_file.is_file():
        raise FileNotFoundError(f"Image not found: {image_file}")

    reader = easyocr.Reader(["en"], gpu=False)
    detections = reader.readtext(str(image_file), detail=1)

    if not detections:
        print("No text detected in the image.")
        return

    items = []
    for box, text, _confidence in detections:
        x_center = sum(point[0] for point in box) / len(box)
        y_center = sum(point[1] for point in box) / len(box)
        height = max(point[1] for point in box) - min(point[1] for point in box)
        items.append((x_center, y_center, height, text))

    headers = [
        "Name of student",
        "Marks in Statistics",
        "Marks in Economics",
        "Marks in Accountancy",
        "Marks in Commerce",
    ]
    column_edges = [0, 0.30, 0.465, 0.625, 0.785, 1]
    with Image.open(image_file) as image:
        image_width = image.width

    row_threshold = median(item[2] for item in items) * 1.25
    rows = []
    for item in sorted(items, key=lambda value: value[1]):
        if not rows or item[1] - rows[-1]["last_y"] > row_threshold:
            rows.append({"last_y": item[1], "items": []})
        rows[-1]["items"].append(item)
        rows[-1]["last_y"] = item[1]

    table = []
    for row in rows[1:]:
        cells = [[] for _ in headers]
        for x_center, _y_center, _height, text in row["items"]:
            column = next(
                index
                for index in range(len(headers))
                if x_center < column_edges[index + 1] * image_width
            )
            cells[column].append((x_center, text))
        table.append([
            " ".join(text for _x, text in sorted(cell))
            for cell in cells
        ])

    df = pd.DataFrame(table, columns=headers)

    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted Table")
        worksheet = writer.sheets["Extracted Table"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        for column in worksheet.columns:
            worksheet.column_dimensions[column[0].column_letter].width = 24
            for cell in column:
                alignment = copy(cell.alignment)
                alignment.wrap_text = True
                alignment.vertical = "top"
                cell.alignment = alignment
        for cell in worksheet[1]:
            font = copy(cell.font)
            font.bold = True
            cell.font = font
        worksheet.row_dimensions[1].height = 36

    print(f"Saved {len(df)} table rows to '{output_excel}'.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Extract English text from an image and save it to Excel."
    )
    parser.add_argument(
        "image",
        nargs="?",
        default="note.jpg",
        help="Image path (default: note.jpg)",
    )
    parser.add_argument(
        "output",
        nargs="?",
        default="student_marks.xlsx",
        help="Excel output path (default: student_marks.xlsx)",
    )
    args = parser.parse_args()
    image_to_excel(args.image, args.output)