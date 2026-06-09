"""Encoding repair and PowerShell output decoding utilities."""

import re
import unicodedata


def decode_powershell_output(raw_output: bytes) -> str:
    """Decodifica stdout do Windows PowerShell 5.1 preservando acentuação."""
    utf8_text = raw_output.decode('utf-8', errors='replace')
    cp1252_text = raw_output.decode('cp1252', errors='replace')
    repaired_cp1252 = repair_mojibake(cp1252_text)

    utf8_score = score_decoded_text(utf8_text)
    repaired_cp1252_score = score_decoded_text(repaired_cp1252)

    if repaired_cp1252_score < utf8_score:
        return repaired_cp1252
    return utf8_text


def repair_mojibake(text: str) -> str:
    """Corrige trechos UTF-8 lidos como cp1252 sem afetar texto já correto."""
    previous_text = text
    for _ in range(3):
        repaired_text = repair_mojibake_once(previous_text)
        if score_decoded_text(repaired_text) >= score_decoded_text(previous_text):
            return previous_text
        previous_text = repaired_text
    return previous_text


def repair_mojibake_once(text: str) -> str:
    """Repara uma iteração de mojibake."""
    parts = re.split(r'(\s+)', text)
    repaired_parts = []
    for part in parts:
        if not part or part.isspace() or not has_mojibake_markers(part):
            repaired_parts.append(part)
            continue

        try:
            repaired_candidate = part.encode('cp1252', errors='strict').decode('utf-8', errors='strict')
        except (UnicodeEncodeError, UnicodeDecodeError):
            repaired_parts.append(part)
            continue

        if score_decoded_text(repaired_candidate) <= score_decoded_text(part):
            repaired_parts.append(repaired_candidate)
        else:
            repaired_parts.append(part)

    return ''.join(repaired_parts)


def has_mojibake_markers(text: str) -> bool:
    """Verifica se text contém marcadores de mojibake."""
    return any(marker in text for marker in ('Ã', 'Â', 'â'))


def score_decoded_text(text: str) -> int:
    """Pontua qualidade de decodificação (menor é melhor)."""
    replacement_penalty = text.count('�') * 10
    mojibake_penalty = sum(text.count(marker) for marker in ('Ã', 'Â', 'â')) * 6
    return replacement_penalty + mojibake_penalty
