"""Reusable workflow helpers for Chinese graduate proposal drafting.

The public API intentionally stays small:

- rewrite_and_balance_proposal(...)
- format_academic_docx(...)
- generate_ppt_blueprint(...)
"""

from .kaiti_workflow import (
    format_academic_docx,
    generate_ppt_blueprint,
    rewrite_and_balance_proposal,
)

__all__ = [
    "rewrite_and_balance_proposal",
    "format_academic_docx",
    "generate_ppt_blueprint",
]
