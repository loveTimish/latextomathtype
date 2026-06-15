# XSC Reference WMF Inventory

Initial inventory built from the uploaded XSC reference ZIP.

## Summary

- DOCX files inspected: 16
- Embedded Word objects: 7851
- WMF preview objects: 7848
- Vector-text WMF objects: 7842
- Bitmap WMF objects: 0
- StretchDIB objects: 0
- Valid Equation.DSMT OLE objects: 7843
- WMF width range: 4.989pt .. 465.000pt
- WMF height range: 6.979pt .. 86.513pt

## Per document

| source | objects | wmf | png | emf | embeddings | Equation.DSMT |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 3-2-6 变速问题 教师版.docx | 573 | 566 | 1 | 5 | 573 | 573 |
| 3-2-6 变速问题 学生版.docx | 182 | 182 | 1 | 5 | 182 | 182 |
| 4-1-1几何图形的认识.教师版.docx | 84 | 82 | 3 | 53 | 84 | 84 |
| 4-3-1 等高模型 教师版.docx | 1556 | 1495 | 4 | 161 | 1556 | 1554 |
| 4-3-1 等高模型 学生版.docx | 435 | 435 | 2 | 87 | 436 | 433 |
| 4-3-3 相似模型.docx | 811 | 772 | 6 | 78 | 811 | 811 |
| 4-3-4 蝴蝶模型 教师版.docx | 527 | 517 | 2 | 53 | 527 | 525 |
| 4-3-4 蝴蝶模型 学生版.docx | 169 | 168 | 2 | 32 | 169 | 168 |
| 4-3-4 鸟头模型 教师版.docx | 388 | 356 | 3 | 30 | 388 | 388 |
| 4-3-4 鸟头模型 学生版.docx | 119 | 119 | 2 | 19 | 119 | 119 |
| 4-3-5 相似模型教师版.docx | 877 | 876 | 4 | 75 | 877 | 877 |
| 4-3-5 相似模型学生版.docx | 253 | 253 | 2 | 47 | 253 | 253 |
| 4-3-6 燕尾定理 教师版.docx | 811 | 772 | 6 | 78 | 812 | 811 |
| 4-3-6 燕尾定理 学生版.docx | 261 | 238 | 2 | 39 | 261 | 261 |
| 8.24课堂练习.docx | 0 | 0 | 2 | 0 | 0 | 0 |
| 8-5 抽屉原理.docx | 805 | 832 | 4 | 16 | 805 | 804 |

## Use in calibration

These documents provide a high-value first calibration set because their previews are overwhelmingly self-contained WMF images and do not rely on bitmap `StretchDIB` records. The immediate gate is to extract source object metrics, generate matching documents, and compare WMF physical size, VML shape size, baseline position, vector-record safety, and MathType OLE validity before tuning `VectorWmfFormulaRenderer` constants.
