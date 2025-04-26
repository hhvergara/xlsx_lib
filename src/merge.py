import os
from pathlib import Path
import aspose.cells as cells

__BASE_DIR__ = Path(__file__).resolve().parent
file_1 = os.path.join(__BASE_DIR__,'examples', 'Example C.xlsx')
file_2 = os.path.join(__BASE_DIR__,'examples', 'Example D.xlsx')


# Load the first Excel file
book1 = cells.Workbook(file_1)

# Load the second Excel file
book2 = cells.Workbook(file_2)

# Merge Files
book1.combine(book2)

# Save Merged File
book1.save("merged-aspose-img-chart.xlsx")