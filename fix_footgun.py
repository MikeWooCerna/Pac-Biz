"""One-shot fix: replace all df.get(col, pd.Series(dtype=float)) footguns
with the safe missing-column pattern across dashboard.py."""

with open('dashboard.py', 'r', encoding='utf-8') as f:
    content = f.read()

replacements = [
    (
        'df.get("score_ai", pd.Series(dtype=float))',
        '(df["score_ai"] if "score_ai" in df.columns else pd.Series(0, index=df.index, dtype=float))'
    ),
    (
        'df.get("score_human", pd.Series(dtype=float))',
        '(df["score_human"] if "score_human" in df.columns else pd.Series(0, index=df.index, dtype=float))'
    ),
    (
        'df.get(ai_col,  pd.Series(dtype=float))',
        '(df[ai_col]  if ai_col  in df.columns else pd.Series(0, index=df.index, dtype=float))'
    ),
    (
        'df.get(ai_col, pd.Series(dtype=float))',
        '(df[ai_col] if ai_col in df.columns else pd.Series(0, index=df.index, dtype=float))'
    ),
    (
        'df.get(max_col, pd.Series(dtype=float))',
        '(df[max_col] if max_col in df.columns else pd.Series(0, index=df.index, dtype=float))'
    ),
]

total = 0
for old, new in replacements:
    n = content.count(old)
    content = content.replace(old, new)
    print(f'  {n:2d}x  {old[:60]}')
    total += n

with open('dashboard.py', 'w', encoding='utf-8') as f:
    f.write(content)

print(f'\nDone — {total} replacements made.')
