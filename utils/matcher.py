from difflib import SequenceMatcher


def flexible_match(name: str, clients_dict: dict) -> tuple:
    """
    Match client name to clients dictionary.
    Returns (account_number, ratio, matched_name) or (None, 0, None).
    """
    name_clean = str(name).strip()
    name_lower = name_clean.lower()
    name_words = name_lower.split()

    # Exact match
    if name_clean in clients_dict:
        return clients_dict[name_clean], 1.0, name_clean

    best_match      = None
    best_ratio      = 0
    best_client_name = None

    for client_name, account_num in clients_dict.items():
        client_lower = client_name.lower()

        # Case-insensitive exact
        if name_lower == client_lower:
            return account_num, 1.0, client_name

        # All words of query appear in client name
        if len(name_words) > 1 and all(w in client_lower for w in name_words):
            ratio = 0.9
            if ratio > best_ratio:
                best_ratio      = ratio
                best_match      = account_num
                best_client_name = client_name
            continue

        # Fuzzy similarity
        ratio = SequenceMatcher(None, name_lower, client_lower).ratio()
        if ratio >= 0.75 and ratio > best_ratio:
            best_ratio      = ratio
            best_match      = account_num
            best_client_name = client_name

    return best_match, best_ratio, best_client_name
