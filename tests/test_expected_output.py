def test_actual_matches_expected(workbook_diffs):
    unexpected = [d for d in workbook_diffs if not d.allowed]
    assert not unexpected, (
        "Actual workbook differs from expected workbook. "
        "See reports/diff_report.html and reports/diff_report.json."
    )
