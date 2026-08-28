def update_match_suggestions(result, edited_display, selected_status):
    """
    Editable Match Suggestion column se changes ko original reconciliation
    result mein wapas save karta hai.
    """
    if edited_display.empty:
        return result

    result = result.copy()

    # Selected status ke rows ke original indexes
    status_mask = result["Status"] == selected_status
    status_indexes = result.index[status_mask].tolist()

    # Edited table ki rows ko original result se map karna
    if len(status_indexes) != len(edited_display):
        return result

    for i, original_idx in enumerate(status_indexes):
        if "Match Suggestion" in edited_display.columns:
            value = edited_display.iloc[i]["Match Suggestion"]

            if pd.isna(value):
                value = ""

            result.at[original_idx, "Match Suggestion"] = str(value)

    return result
