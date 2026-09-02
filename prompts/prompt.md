Inspect samples/input.xlsx and samples/expected_output.xlsx and infer the exact
transformation between them.
The intended business changes are:

- [column with header in row 9 "Сумма годовой корректировки" should be moved and positioned to the right after column with name "Сумма годовой корректировки". The formula in the column should show the same result so if needed adjust references accordingly]
- [column with header in row 9 "Перерасчет суммы годовой корректировки" should be moved and positioned to the left before column with name "Оплачено нарастающим итогом с начала периода". The formula in the column should show the same result so if needed adjust references accordingly]
- [ column with header in row 9 "Дата заключения" 
  is displayed in different formats in different cells. All the cells should show data only in the format "ДД.ММ.ГГГГ".]
- [add the formula calculating the sum of the column values to the cells in row 7 and columns BB and CV]
  
- Treat all other differences as accidental and ask me about any ambiguity,
  especially formula-reference behavior. Identify columns by headers rather than
  fixed Excel letters. Preserve native Excel data types and unrelated workbook
  content. Do not modify either reference file.
