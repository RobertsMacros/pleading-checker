#!/usr/bin/env python3
"""
Fix all VBA compile errors across the pleading-checker project.

Three classes of issues:
1. Scripting.Dictionary early binding -> late-bound Object + CreateObject
2. PleadingsIssue early binding -> late-bound Object + Dictionary-based issues
3. Non-ASCII characters in comments -> plain ASCII

Also converts direct PleadingsEngine.Xxx calls in rule files to
Application.Run so rules compile standalone.
"""

import os
import re
import glob


def fix_non_ascii(text):
    """Replace non-ASCII box-drawing and decorative chars with ASCII equivalents."""
    replacements = {
        '\u2550': '=',   # double horizontal
        '\u2500': '-',   # light horizontal
        '\u2014': '--',  # em dash
        '\u2192': '->',  # rightwards arrow
        '\u2194': '<->', # left right arrow
        '\u2212': '-',   # minus sign
        '\u00A0': ' ',   # non-breaking space
        '\u2013': '-',   # en dash
    }

    lines = text.split('\n')
    result = []
    for line in lines:
        stripped = line.lstrip()
        is_comment = stripped.startswith("'")

        if is_comment:
            for old, new in replacements.items():
                line = line.replace(old, new)
            # Replace any remaining non-ASCII in comments
            cleaned = []
            for ch in line:
                if ord(ch) > 127:
                    cleaned.append('?')
                else:
                    cleaned.append(ch)
            line = ''.join(cleaned)
        else:
            # In code lines, replace all known non-ASCII
            for old, new in replacements.items():
                line = line.replace(old, new)
            # Also replace smart quotes in code (shouldn't be there)
            line = line.replace('\u2018', "'")
            line = line.replace('\u2019', "'")
            line = line.replace('\u201C', '"')
            line = line.replace('\u201D', '"')
        result.append(line)
    return '\n'.join(result)


def fix_scripting_dictionary(text):
    """Convert all Scripting.Dictionary early binding to late binding."""

    # 1. Fix "Dim x As New Scripting.Dictionary" -> two lines
    def fix_dim_as_new_sd(match):
        indent = match.group(1)
        varname = match.group(2)
        comment = match.group(3) or ''
        return (f'{indent}Dim {varname} As Object{comment}\n'
                f'{indent}Set {varname} = CreateObject("Scripting.Dictionary")')

    text = re.sub(
        r'^([ \t]*)Dim (\w+) As New Scripting\.Dictionary([ \t]*\'.*)?$',
        fix_dim_as_new_sd,
        text,
        flags=re.MULTILINE
    )

    # 2. Fix all remaining "As Scripting.Dictionary" (params, return types, Dim)
    text = re.sub(
        r'\bAs Scripting\.Dictionary\b',
        'As Object',
        text
    )

    # 3. Fix "Set x = New Scripting.Dictionary" -> CreateObject
    text = re.sub(
        r'(Set\s+\w+\s*=\s*)New Scripting\.Dictionary',
        r'\1CreateObject("Scripting.Dictionary")',
        text
    )

    # 4. Fix remaining "New Scripting.Dictionary" in expressions
    text = re.sub(
        r'\bNew Scripting\.Dictionary\b',
        'CreateObject("Scripting.Dictionary")',
        text
    )

    return text


def fix_pleadings_issue(text, filename):
    """
    Convert PleadingsIssue references to Dictionary-based issues.
    """
    if 'PleadingsIssue' not in text:
        return text

    base = os.path.basename(filename)
    if base == 'PleadingsIssue.cls':
        return text

    # 1. "Dim x As New PleadingsIssue" -> "Dim x As Object"
    #    (the New will be handled when we replace the .Init call)
    def fix_dim_as_new_pi(match):
        indent = match.group(1)
        varname = match.group(2)
        return f'{indent}Dim {varname} As Object'

    text = re.sub(
        r'^([ \t]*)Dim (\w+) As New PleadingsIssue\b.*$',
        fix_dim_as_new_pi,
        text,
        flags=re.MULTILINE
    )

    # 2. "Dim x As PleadingsIssue" -> "Dim x As Object"
    text = re.sub(
        r'(\bDim\s+\w+\s+)As PleadingsIssue\b',
        r'\1As Object',
        text
    )

    # 3. Remove standalone "Set x = New PleadingsIssue" lines
    text = re.sub(
        r'^[ \t]*Set \w+ = New PleadingsIssue\s*\n',
        '',
        text,
        flags=re.MULTILINE
    )

    # 4. Replace "x.Init arg1, arg2, ..." with "Set x = CreateIssueDict(arg1, arg2, ...)"
    #    Handle multi-line with _ continuation
    def replace_init(match):
        indent = match.group(1)
        varname = match.group(2)
        args_raw = match.group(3)
        # Remove line continuations and join
        args = re.sub(r'\s*_\s*\n\s*', ' ', args_raw)
        args = args.strip()
        return f'{indent}Set {varname} = CreateIssueDict({args})'

    text = re.sub(
        r'^([ \t]*)(\w+)\.Init\s+((?:.*?_\s*\n\s*)*.*?)$',
        replace_init,
        text,
        flags=re.MULTILINE
    )

    # 5. Add CreateIssueDict helper if needed and not present
    if 'CreateIssueDict(' in text and 'Private Function CreateIssueDict' not in text:
        helper = '''
' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based issue (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function
'''
        text = text.rstrip() + '\n' + helper

    return text


def fix_engine_calls_in_rules(text, filename):
    """
    Convert direct PleadingsEngine.Xxx calls to Application.Run in rule files.
    """
    base = os.path.basename(filename)
    if not base.startswith('Rules_') and not base.startswith('Rule'):
        return text

    needs_page_range = 'PleadingsEngine.IsInPageRange' in text
    needs_location = 'PleadingsEngine.GetLocationString' in text
    needs_whitelist = 'PleadingsEngine.IsWhitelistedTerm' in text
    needs_spelling_mode = 'PleadingsEngine.GetSpellingMode' in text
    needs_set_page_range = 'PleadingsEngine.SetPageRange' in text
    needs_set_whitelist = 'PleadingsEngine.SetWhitelist' in text

    if needs_page_range:
        text = text.replace('PleadingsEngine.IsInPageRange(', 'EngineIsInPageRange(')
        text = text.replace('PleadingsEngine.IsInPageRange', 'EngineIsInPageRange')

    if needs_location:
        text = text.replace('PleadingsEngine.GetLocationString(', 'EngineGetLocationString(')
        text = text.replace('PleadingsEngine.GetLocationString', 'EngineGetLocationString')

    if needs_whitelist:
        text = text.replace('PleadingsEngine.IsWhitelistedTerm(', 'EngineIsWhitelistedTerm(')

    if needs_spelling_mode:
        text = text.replace('PleadingsEngine.GetSpellingMode()', 'EngineGetSpellingMode()')

    if needs_set_page_range:
        text = re.sub(
            r'PleadingsEngine\.SetPageRange\s+(\S+),\s*(\S+)',
            r'EngineSetPageRange \1, \2',
            text
        )

    if needs_set_whitelist:
        text = re.sub(
            r'PleadingsEngine\.SetWhitelist\s+(\S+)',
            r'EngineSetWhitelist \1',
            text
        )

    # Build helper functions
    helpers = []

    if needs_page_range:
        helpers.append('''\
' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function''')

    if needs_location:
        helpers.append('''\
' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function''')

    if needs_whitelist:
        helpers.append('''\
' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsWhitelistedTerm
' ----------------------------------------------------------------
Private Function EngineIsWhitelistedTerm(ByVal term As String) As Boolean
    On Error Resume Next
    EngineIsWhitelistedTerm = Application.Run("PleadingsEngine.IsWhitelistedTerm", term)
    If Err.Number <> 0 Then
        EngineIsWhitelistedTerm = False
        Err.Clear
    End If
    On Error GoTo 0
End Function''')

    if needs_spelling_mode:
        helpers.append('''\
' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetSpellingMode
' ----------------------------------------------------------------
Private Function EngineGetSpellingMode() As String
    On Error Resume Next
    EngineGetSpellingMode = Application.Run("PleadingsEngine.GetSpellingMode")
    If Err.Number <> 0 Then
        EngineGetSpellingMode = "UK"
        Err.Clear
    End If
    On Error GoTo 0
End Function''')

    if needs_set_page_range:
        helpers.append('''\
' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.SetPageRange
' ----------------------------------------------------------------
Private Sub EngineSetPageRange(ByVal startPg As Long, ByVal endPg As Long)
    On Error Resume Next
    Application.Run "PleadingsEngine.SetPageRange", startPg, endPg
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub''')

    if needs_set_whitelist:
        helpers.append('''\
' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.SetWhitelist
' ----------------------------------------------------------------
Private Sub EngineSetWhitelist(dict As Object)
    On Error Resume Next
    Application.Run "PleadingsEngine.SetWhitelist", dict
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub''')

    if helpers:
        text = text.rstrip() + '\n\n' + '\n\n'.join(helpers) + '\n'

    return text


def fix_engine_issue_access(text):
    """
    In PleadingsEngine.bas, update issue property access to support
    Dictionary-based issues alongside PleadingsIssue objects.
    Also add CreateIssue factory and JSON helpers.
    """
    # Replace direct property reads on issues with GetIssueProp()
    # We need to be careful to only replace in the right context

    # Replace .ToJSON calls
    text = re.sub(r'(\w+)\.ToJSON\b(?!\s*=)', r'IssueToJSON(\1)', text)

    # Replace property reads on issue objects
    # Pattern: varname.PropertyName (not followed by = which is a write)
    for prop in ['RuleName', 'Location', 'Issue', 'Suggestion', 'Severity',
                 'RangeStart', 'RangeEnd', 'AutoFixSafe']:
        # But NOT in Property Get/Let definitions, NOT in class code
        # Only match when preceded by a variable name (not by "Property Get/Let")
        text = re.sub(
            r'(?<!Property Get )(?<!Property Let )(\b(?:issue|iss|item|curIssue)\b)\.'
            + prop + r'\b(?!\s*=)',
            rf'GetIssueProp(\1, "{prop}")',
            text,
            flags=re.IGNORECASE
        )

    # Add helper functions if not present
    if 'Private Function GetIssueProp(' not in text:
        helpers = '''
' ================================================================
'  PRIVATE: Read a property from an issue (Dict or PleadingsIssue)
' ================================================================
Private Function GetIssueProp(iss As Object, ByVal propName As String) As Variant
    On Error Resume Next
    If TypeName(iss) = "Dictionary" Then
        GetIssueProp = iss(propName)
    Else
        CallByName iss, propName, VbGet
        GetIssueProp = CallByName(iss, propName, VbGet)
    End If
    If Err.Number <> 0 Then
        GetIssueProp = ""
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ================================================================
'  PRIVATE: Format an issue as JSON (Dict or PleadingsIssue)
' ================================================================
Private Function IssueToJSON(iss As Object) As String
    Dim s As String
    s = "    {" & vbCrLf
    s = s & "      ""rule"": """ & EscJSON(CStr(GetIssueProp(iss, "RuleName"))) & """," & vbCrLf
    s = s & "      ""location"": """ & EscJSON(CStr(GetIssueProp(iss, "Location"))) & """," & vbCrLf
    s = s & "      ""severity"": """ & EscJSON(CStr(GetIssueProp(iss, "Severity"))) & """," & vbCrLf
    s = s & "      ""issue"": """ & EscJSON(CStr(GetIssueProp(iss, "Issue"))) & """," & vbCrLf
    s = s & "      ""suggestion"": """ & EscJSON(CStr(GetIssueProp(iss, "Suggestion"))) & """," & vbCrLf
    s = s & "      ""auto_fix_safe"": " & IIf(CBool("0" & GetIssueProp(iss, "AutoFixSafe")), "true", "false") & vbCrLf
    s = s & "    }"
    IssueToJSON = s
End Function

Private Function EscJSON(ByVal txt As String) As String
    txt = Replace(txt, "\\", "\\\\")
    txt = Replace(txt, """", "\\""")
    txt = Replace(txt, vbCr, "\\r")
    txt = Replace(txt, vbLf, "\\n")
    txt = Replace(txt, vbTab, "\\t")
    EscJSON = txt
End Function

' ================================================================
'  PUBLIC: Factory to create a dictionary-based issue
' ================================================================
Public Function CreateIssue(ByVal ruleName_ As String, _
                            ByVal location_ As String, _
                            ByVal issue_ As String, _
                            ByVal suggestion_ As String, _
                            ByVal rangeStart_ As Long, _
                            ByVal rangeEnd_ As Long, _
                            Optional ByVal severity_ As String = "error", _
                            Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssue = d
End Function
'''
        text = text.rstrip() + '\n' + helpers

    return text


def update_dependency_comments(text):
    """Remove Microsoft Scripting Runtime dependency comments."""
    text = re.sub(
        r"^[ \t]*'[ \t]*-[ \t]*Microsoft Scripting Runtime.*$\n?",
        '',
        text,
        flags=re.MULTILINE
    )
    # Remove PleadingsIssue dependency comments in rule files
    text = re.sub(
        r"^[ \t]*'[ \t]*-[ \t]*PleadingsIssue\.cls.*$\n?",
        '',
        text,
        flags=re.MULTILINE
    )
    return text


def fix_previous_run_bugs(text):
    """Fix artifacts from previous buggy script runs."""

    # Fix "Set Dim x = ..." -> split into Dim + Set
    def fix_set_dim(match):
        indent = match.group(1)
        varname = match.group(2)
        rest = match.group(3)
        return f'{indent}Dim {varname} As Object\n{indent}Set {varname}{rest}'

    text = re.sub(
        r'^([ \t]*)Set Dim (\w+)( = .*)$',
        fix_set_dim,
        text,
        flags=re.MULTILINE
    )

    # Fix "Dim Dim x As Object" -> "Dim x As Object"
    text = re.sub(
        r'\bDim Dim (\w+)',
        r'Dim \1',
        text
    )

    # Remove duplicate adjacent "Dim x As Object" lines
    lines = text.split('\n')
    result = []
    prev_line_stripped = ''
    for line in lines:
        stripped = line.strip()
        if stripped == prev_line_stripped and stripped.startswith('Dim ') and stripped.endswith('As Object'):
            continue  # skip duplicate
        result.append(line)
        prev_line_stripped = stripped
    text = '\n'.join(result)

    return text


def process_file(filepath):
    """Process a single VBA file."""
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        original = f.read()

    text = original
    basename = os.path.basename(filepath)

    # 0. Fix bugs from previous run
    text = fix_previous_run_bugs(text)

    # 1. Fix non-ASCII
    text = fix_non_ascii(text)

    # 2. Fix Scripting.Dictionary
    text = fix_scripting_dictionary(text)

    # 3. Fix PleadingsIssue (skip the class itself)
    text = fix_pleadings_issue(text, filepath)

    # 4. Fix PleadingsEngine.Xxx calls (rule files only)
    if basename.startswith('Rules_') or basename.startswith('Rule'):
        text = fix_engine_calls_in_rules(text, filepath)

    # 5. Special engine handling
    if basename == 'PleadingsEngine.bas':
        text = fix_engine_issue_access(text)

    # 6. Launcher fix
    if basename == 'PleadingsLauncher.bas':
        text = re.sub(r'(\bDim\s+\w+\s+)As Scripting\.Dictionary', r'\1As Object', text)

    # 7. Cleanup dependency comments
    text = update_dependency_comments(text)

    # 8. Remove duplicate helper functions (from multiple runs)
    # Count occurrences of CreateIssueDict function
    count = text.count('Private Function CreateIssueDict(')
    if count > 1:
        # Keep only the last one
        parts = text.split('Private Function CreateIssueDict(')
        text = parts[0]
        for i in range(1, len(parts) - 1):
            # Skip this duplicate - find the End Function and remove
            end_idx = parts[i].find('End Function')
            if end_idx >= 0:
                text += parts[i][end_idx + len('End Function') + 1:]
        text += 'Private Function CreateIssueDict(' + parts[-1]

    # Same for engine helpers
    for func_name in ['EngineIsInPageRange', 'EngineGetLocationString',
                      'EngineIsWhitelistedTerm', 'EngineGetSpellingMode',
                      'EngineSetPageRange', 'EngineSetWhitelist']:
        marker = f'Private Function {func_name}('
        if func_name.startswith('EngineSet'):
            marker = f'Private Sub {func_name}('
        count = text.count(marker)
        if count > 1:
            parts = text.split(marker)
            text = parts[0]
            for i in range(1, len(parts) - 1):
                end_marker = 'End Function' if 'Function' in marker else 'End Sub'
                end_idx = parts[i].find(end_marker)
                if end_idx >= 0:
                    text += parts[i][end_idx + len(end_marker) + 1:]
            text += marker + parts[-1]

    if text != original:
        with open(filepath, 'w', encoding='utf-8', newline='\r\n') as f:
            f.write(text)
        return True
    return False


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))

    patterns = ['*.bas', '*.cls', '*.frm']
    changed_files = []

    # Process Combined/
    combined_dir = os.path.join(base_dir, 'Combined')
    for pattern in patterns:
        for filepath in sorted(glob.glob(os.path.join(combined_dir, pattern))):
            print(f"Processing: Combined/{os.path.basename(filepath)}")
            if process_file(filepath):
                changed_files.append(filepath)
                print(f"  -> FIXED")
            else:
                print(f"  -> OK")

    # Process root
    for pattern in patterns:
        for filepath in sorted(glob.glob(os.path.join(base_dir, pattern))):
            print(f"Processing: {os.path.basename(filepath)}")
            if process_file(filepath):
                changed_files.append(filepath)
                print(f"  -> FIXED")
            else:
                print(f"  -> OK")

    # Process Rules/
    rules_dir = os.path.join(base_dir, 'Rules')
    if os.path.isdir(rules_dir):
        for pattern in patterns:
            for filepath in sorted(glob.glob(os.path.join(rules_dir, pattern))):
                print(f"Processing: Rules/{os.path.basename(filepath)}")
                if process_file(filepath):
                    changed_files.append(filepath)
                    print(f"  -> FIXED")
                else:
                    print(f"  -> OK")

    print(f"\n{'='*60}")
    print(f"Fixed {len(changed_files)} files.")

    # Verification pass
    print(f"\n{'='*60}")
    print("VERIFICATION: Checking for remaining issues...")

    all_files = []
    for d in [combined_dir, base_dir, rules_dir]:
        if os.path.isdir(d):
            for pattern in patterns:
                all_files.extend(glob.glob(os.path.join(d, pattern)))

    issues_found = 0
    for filepath in all_files:
        with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        basename = os.path.basename(filepath)

        if basename == 'PleadingsIssue.cls':
            continue

        # Check for remaining early binding
        if re.search(r'\bAs Scripting\.Dictionary\b', content):
            print(f"  WARN: {basename} still has 'As Scripting.Dictionary'")
            issues_found += 1

        if re.search(r'\bNew Scripting\.Dictionary\b', content):
            print(f"  WARN: {basename} still has 'New Scripting.Dictionary'")
            issues_found += 1

        if re.search(r'\bAs PleadingsIssue\b', content):
            print(f"  WARN: {basename} still has 'As PleadingsIssue'")
            issues_found += 1

        if re.search(r'\bNew PleadingsIssue\b', content):
            print(f"  WARN: {basename} still has 'New PleadingsIssue'")
            issues_found += 1

        if re.search(r'^[ \t]*Set Dim ', content, re.MULTILINE):
            print(f"  WARN: {basename} has invalid 'Set Dim' syntax")
            issues_found += 1

        # Check for non-ASCII (skip .py files)
        if not filepath.endswith('.py'):
            for i, line in enumerate(content.split('\n'), 1):
                if any(ord(c) > 127 for c in line):
                    # Only flag if not in a string literal
                    stripped = line.lstrip()
                    if stripped.startswith("'") or not ('"' in line):
                        print(f"  WARN: {basename}:{i} has non-ASCII chars")
                        issues_found += 1
                        break

        # Check for direct PleadingsEngine. calls in rule files
        if basename.startswith('Rules_') or basename.startswith('Rule'):
            # Skip comments
            for i, line in enumerate(content.split('\n'), 1):
                stripped = line.lstrip()
                if stripped.startswith("'"):
                    continue
                if 'PleadingsEngine.' in stripped:
                    print(f"  WARN: {basename}:{i} has direct PleadingsEngine. call")
                    issues_found += 1
                    break

    if issues_found == 0:
        print("  ALL CLEAR - no compile risks found!")
    else:
        print(f"  {issues_found} issue(s) remain")


if __name__ == '__main__':
    main()
