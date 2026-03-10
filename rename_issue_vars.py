#!/usr/bin/env python3
"""Rename issue/issueX variables to finding/findingX in rule modules."""
import os
import re
import glob

def rename_issue_vars(text):
    """
    Rename PleadingsIssue-derived variables:
      issue -> finding, issueD -> findingD, issueA -> findingA, etc.

    Preserve:
      - issues (Collection name)
      - issueText, issueStr (string variables)
      - "Issue" (dictionary keys in quotes)
      - CreateIssueDict (function name)
      - summaryIssue -> summaryFinding
    """
    # Map of old -> new for specific variant names found in the codebase
    # We handle the generic 'issue' separately with word boundary

    # First handle specific multi-word variants that contain 'issue'
    specific_renames = [
        ('summaryIssue', 'summaryFinding'),
        ('issueUnused', 'findingUnused'),
        ('issuePara', 'findingPara'),
        ('issueRun', 'findingRun'),
        ('issueFN', 'findingFN'),
        ('issueLC', 'findingLC'),
        ('issueH', 'findingH'),
        ('issueA', 'findingA'),
        ('issueB', 'findingB'),
        ('issueD', 'findingD'),
        ('issueT', 'findingT'),
        ('issue33', 'finding33'),
    ]

    for old, new in specific_renames:
        text = re.sub(r'\b' + old + r'\b', new, text)

    # Now handle standalone 'issue' (but NOT issues, issueText, issueStr, issueIdx, etc.)
    # Pattern: 'issue' at word boundary, NOT followed by [a-z] (which would make it issues, issueText, etc.)
    # BUT we already renamed issueA, issueD etc. above, so those are gone
    # We need: \bissue\b but not \bissues\b or \bissueText\b
    text = re.sub(r'\bissue\b(?!s\b|Text|Str|Idx|Count|Num|_)', 'finding', text)

    return text


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    combined_dir = os.path.join(base_dir, 'Combined')

    changed = []
    for filepath in sorted(glob.glob(os.path.join(combined_dir, 'Rules_*.bas'))):
        with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
            original = f.read()

        text = rename_issue_vars(original)

        if text != original:
            with open(filepath, 'w', encoding='utf-8', newline='\r\n') as f:
                f.write(text)
            changed.append(os.path.basename(filepath))
            print(f"  RENAMED: {os.path.basename(filepath)}")

    # Also fix PleadingsEngine.bas (issue vars there too)
    engine_path = os.path.join(combined_dir, 'PleadingsEngine.bas')
    with open(engine_path, 'r', encoding='utf-8', errors='replace') as f:
        original = f.read()
    text = rename_issue_vars(original)
    if text != original:
        with open(engine_path, 'w', encoding='utf-8', newline='\r\n') as f:
            f.write(text)
        changed.append('PleadingsEngine.bas')
        print(f"  RENAMED: PleadingsEngine.bas")

    # Also clean up remaining "PleadingsIssue" references in comments
    for filepath in sorted(glob.glob(os.path.join(combined_dir, '*.bas'))):
        with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
            text = f.read()
        basename = os.path.basename(filepath)

        original = text
        # Remove comment lines mentioning PleadingsIssue as a dependency
        text = re.sub(
            r"^[ \t]*'.*PleadingsIssue\.cls.*$\n?",
            '',
            text,
            flags=re.MULTILINE
        )
        # Replace "PleadingsIssue" in remaining comments with "issue dictionary"
        lines = text.split('\n')
        result = []
        for line in lines:
            if line.lstrip().startswith("'"):
                line = line.replace('PleadingsIssue', 'issue dictionary')
            result.append(line)
        text = '\n'.join(result)

        if text != original:
            with open(filepath, 'w', encoding='utf-8', newline='\r\n') as f:
                f.write(text)
            if basename not in changed:
                changed.append(basename)
            print(f"  CLEANED: {basename}")

    print(f"\nDone. {len(changed)} files updated.")

    # Verify no PleadingsIssue remains outside PleadingsIssue.cls
    print("\nVerification:")
    for filepath in sorted(glob.glob(os.path.join(combined_dir, '*.bas'))):
        with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
            for i, line in enumerate(f, 1):
                if 'PleadingsIssue' in line and not line.lstrip().startswith("'"):
                    print(f"  WARN: {os.path.basename(filepath)}:{i}: {line.strip()}")


if __name__ == '__main__':
    main()
