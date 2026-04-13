# -*- coding: utf-8 -*-
from modules.planning.application.print_sheet_prep import prepare_sheet_candidates


def test_prepare_sheet_candidates_empty() -> None:
    assert prepare_sheet_candidates([]) == []
