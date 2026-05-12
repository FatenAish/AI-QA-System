"""
=============================================================================
ARABIC WORD-BY-WORD DIFF PATCH
=============================================================================
Replace these 3 functions in your existing app.py.
No other changes needed.

Search for each function name and replace the entire def block.
=============================================================================
"""

# =============================================================================
# REPLACE: _explode_change_to_micro_edits
# =============================================================================
#
# What changed:
#   - REMOVED the conservative threshold that was blocking splits when
#     token count < 10 or non_equal_estimate < 3 (was hiding small Arabic edits)
#   - delete / insert blocks are now also word-split (previously only replace was)
#     so removing one word = 1 row, removing 5 words = 5 rows
#   - fallback to [change] only when tokenisation is genuinely empty
#
def _explode_change_to_micro_edits(change, lang):
    original   = change.get("original", "") or ""
    revised    = change.get("revised", "") or ""
    tag        = change.get("tag", "")
    has_arabic = bool(re.search(r"[\u0600-\u06FF]", original + revised))

    # ── delete / insert: split into one row per word/token ─────────────────
    if tag in ("delete", "insert") and (has_arabic or lang == "Arabic"):
        text   = original if tag == "delete" else revised
        tokens = _tokenize_for_micro_diff(text)
        if len(tokens) < 2:
            return [change]   # single token — keep as-is
        edits = []
        for i, tok in enumerate(tokens):
            if not tok.strip():
                continue
            ctx = _micro_context(tokens, i, i + 1)
            edits.append({
                **{k: v for k, v in change.items()
                   if k not in {"original", "revised", "similarity", "word_delta", "tag"}},
                "tag":              tag,
                "original":         tok if tag == "delete" else "",
                "revised":          "" if tag == "delete" else tok,
                "original_context": ctx if tag == "delete" else "",
                "revised_context":  "" if tag == "delete" else ctx,
                "similarity":       0.0,
                "word_delta":       -1 if tag == "delete" else 1,
                "micro_edit":       True,
            })
        return edits if edits else [change]

    # ── replace: token-level SequenceMatcher ───────────────────────────────
    if tag != "replace":
        return [change]

    old_tokens = _tokenize_for_micro_diff(original)
    new_tokens = _tokenize_for_micro_diff(revised)
    if not old_tokens or not new_tokens:
        return [change]

    # English non-Arabic: keep the old conservative guard
    if not has_arabic and lang != "Arabic":
        non_equal_estimate = abs(len(old_tokens) - len(new_tokens))
        if max(len(old_tokens), len(new_tokens)) < 10 and non_equal_estimate < 3:
            return [change]

    # Arabic: NO size threshold — always split word-by-word
    old_norm = [normalize_for_compare(t) for t in old_tokens]
    new_norm = [normalize_for_compare(t) for t in new_tokens]
    sm       = difflib.SequenceMatcher(None, old_norm, new_norm, autojunk=False)

    edits = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            continue

        old_part = " ".join(old_tokens[i1:i2]).strip()
        new_part = " ".join(new_tokens[j1:j2]).strip()

        if not old_part and not new_part:
            continue
        if looks_like_formatting_only(old_part, new_part):
            continue
        if len(normalize_for_compare(old_part + new_part)) < 2:
            continue

        old_ctx = _micro_context(old_tokens, i1, i2)
        new_ctx = _micro_context(new_tokens, j1, j2)

        edits.append({
            **{k: v for k, v in change.items()
               if k not in {"original", "revised", "similarity", "word_delta", "tag"}},
            "tag":              op,
            "original":         old_part[:700],
            "revised":          new_part[:700],
            "original_context": old_ctx[:700],
            "revised_context":  new_ctx[:700],
            "similarity":       round(token_similarity(old_part, new_part), 3),
            "word_delta":       len(new_part.split()) - len(old_part.split()),
            "micro_edit":       True,
        })

    return edits if edits else [change]


# =============================================================================
# REPLACE: explode_changes_to_micro_edits
# =============================================================================
#
# What changed:
#   - Now passes delete AND insert blocks through the splitter, not only replace.
#     (The original only called the function but the old inner logic skipped
#      non-replace tags; the new inner logic handles them too.)
#
def explode_changes_to_micro_edits(changes, lang):
    exploded = []
    for ch in changes or []:
        exploded.extend(_explode_change_to_micro_edits(ch, lang))
    return exploded


# =============================================================================
# REPLACE: compute_diff
# =============================================================================
#
# What changed:
#   - Added `lang` parameter (default "Arabic") so micro-explosion knows the language.
#   - Micro-explosion now runs BEFORE deduplication, preserving per-word context.
#   - Formatting-only micro rows dropped before returning.
#
# You also need to update every call-site of compute_diff to pass lang:
#
#   In fetch_editor_handoff_revisions call inside page_gdoc_submit():
#       diff_changes = compute_diff(handoff_writer_text, handoff_editor_text, lang)
#
#   In compute_consecutive_revision_diffs():
#       changes = compute_diff(prev_text, cur_text, lang)
#       ...
#       fallback = compute_diff(first_text, last_text, lang)
#
def compute_diff(writer_text, editor_text, lang="Arabic"):
    writer_text = clean_google_doc_export_artifacts(writer_text)
    editor_text = clean_google_doc_export_artifacts(editor_text)

    w_sents = split_sentences_smart(writer_text)
    e_sents = split_sentences_smart(editor_text)

    w_keys = [normalize_for_compare(x) for x in w_sents]
    e_keys = [normalize_for_compare(x) for x in e_sents]

    sm      = difflib.SequenceMatcher(None, w_keys, e_keys, autojunk=False)
    changes = []

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue

        original = sanitize_diff_side(" ".join(w_sents[i1:i2]).strip())
        revised  = sanitize_diff_side(" ".join(e_sents[j1:j2]).strip())

        if not original and not revised:
            continue
        if _looks_like_comment_artifact(original, revised):
            continue
        if not original or not revised:
            continue
        if looks_like_formatting_only(original, revised):
            continue

        sim = round(token_similarity(original, revised), 3)
        changes.append({
            "tag":        tag,
            "original":   original[:700],
            "revised":    revised[:700],
            "similarity": sim,
            "word_delta": len(revised.split()) - len(original.split()),
        })

    # Step 1: for Arabic, split large paragraph blocks at punctuation first
    if lang == "Arabic":
        changes = _split_arabic_large_changes(changes)

    # Step 2: explode ALL blocks (replace / delete / insert) to word level
    changes = explode_changes_to_micro_edits(changes, lang)

    # Step 3: drop formatting-only micro rows (tashkeel, punctuation-only)
    changes = [
        ch for ch in changes
        if not looks_like_formatting_only(
            ch.get("original", ""), ch.get("revised", "")
        )
    ]

    # Step 4: deduplicate
    changes = _dedupe_diff_changes(changes)

    return changes


# =============================================================================
# CALL-SITE UPDATES — find these two blocks in page_gdoc_submit() and update:
# =============================================================================
#
# OLD:
#   diff_changes = compute_diff(handoff_writer_text, handoff_editor_text)
# NEW:
#   diff_changes = compute_diff(handoff_writer_text, handoff_editor_text, lang)
#
# ─────────────────────────────────────────────────────────────────────────────
#
# OLD (inside compute_consecutive_revision_diffs):
#   changes = compute_diff(prev_text, cur_text)
#   ...
#   fallback = compute_diff(first_text, last_text)
# NEW:
#   changes = compute_diff(prev_text, cur_text, lang)
#   ...
#   fallback = compute_diff(first_text, last_text, lang)
#
# =============================================================================
