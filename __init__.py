"""Generic academic proposal workflow helpers.

The package exposes three primary functions:

- synthesize_and_rewrite_proposal(...)
- format_academic_docx(...)
- generate_generic_ppt_blueprint(...)
"""

from .kaiti_workflow import (
    format_academic_docx,
    generate_generic_ppt_blueprint,
    generate_ppt_blueprint,
    rewrite_and_balance_proposal,
    synthesize_and_rewrite_proposal,
)

__all__ = [
    "synthesize_and_rewrite_proposal",
    "format_academic_docx",
    "generate_generic_ppt_blueprint",
    "rewrite_and_balance_proposal",
    "generate_ppt_blueprint",
]
