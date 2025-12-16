"""
PDF Parser for Insurance Quotes and Binders

This module handles parsing of insurance PDF documents to extract:
- Policy limits
- Attachment points (excess of)
- Carrier information
- Premium amounts
- Policy numbers
- Risk Bearing Entities (RBEs)
"""

import re
from typing import Dict, List, Optional, Any
import io


def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    """
    Extract text content from PDF bytes.

    Args:
        pdf_bytes: PDF file as bytes

    Returns:
        Extracted text content
    """
    try:
        import pdfplumber

        text_content = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text_content.append(page_text)

        return "\n".join(text_content)
    except Exception as e:
        raise Exception(f"Error extracting PDF text: {str(e)}")


def parse_currency(value_str: str) -> float:
    """
    Parse currency string to float.
    Examples: "$1,000,000" -> 1000000, "1M" -> 1000000

    Args:
        value_str: Currency string

    Returns:
        Float value
    """
    if not value_str:
        return 0.0

    # Remove currency symbols and whitespace
    clean_str = re.sub(r"[$,\s]", "", value_str.upper())

    # Handle M (millions) and K (thousands) suffixes
    multiplier = 1
    if "M" in clean_str:
        multiplier = 1_000_000
        clean_str = clean_str.replace("M", "")
    elif "K" in clean_str:
        multiplier = 1_000
        clean_str = clean_str.replace("K", "")

    try:
        return float(clean_str) * multiplier
    except ValueError:
        return 0.0


def extract_limit_patterns(text: str) -> List[Dict[str, Any]]:
    """
    Extract insurance limit patterns from text.
    Looks for patterns like:
    - $1,000,000 excess of $1,000,000
    - $5M xs $1M
    - $1,000,000 Primary
    - Limit: $1,000,000

    Args:
        text: Text content to search

    Returns:
        List of extracted limit information
    """
    limits = []

    # Pattern 1: "X excess of Y" or "X xs Y"
    excess_pattern = (
        r"\$?[\d,]+(?:\.\d+)?[MK]?\s*(?:excess\s+of|xs|x/s)\s*\$?[\d,]+(?:\.\d+)?[MK]?"
    )
    excess_matches = re.finditer(excess_pattern, text, re.IGNORECASE)

    for match in excess_matches:
        match_text = match.group()
        # Split on "excess of" or "xs"
        parts = re.split(r"excess\s+of|xs|x/s", match_text, flags=re.IGNORECASE)
        if len(parts) == 2:
            limit = parse_currency(parts[0].strip())
            attachment = parse_currency(parts[1].strip())
            limits.append(
                {
                    "limit": limit,
                    "attachment": attachment,
                    "is_primary": False,
                    "raw_text": match_text.strip(),
                }
            )

    # Pattern 2: "Primary" or "Primary Layer"
    primary_pattern = r"\$?[\d,]+(?:\.\d+)?[MK]?\s*(?:primary|primary\s+layer)"
    primary_matches = re.finditer(primary_pattern, text, re.IGNORECASE)

    for match in primary_matches:
        match_text = match.group()
        limit = parse_currency(match_text)
        if limit > 0:
            limits.append(
                {
                    "limit": limit,
                    "attachment": 0,
                    "is_primary": True,
                    "raw_text": match_text.strip(),
                }
            )

    # Pattern 3: "Limit: $X"
    limit_pattern = r"(?:limit|coverage)[:\s]+\$?[\d,]+(?:\.\d+)?[MK]?"
    limit_matches = re.finditer(limit_pattern, text, re.IGNORECASE)

    for match in limit_matches:
        match_text = match.group()
        limit = parse_currency(match_text.split(":")[-1].strip())
        if limit > 0:
            # Check if there's an attachment mentioned nearby
            context = text[
                max(0, match.start() - 100) : min(len(text), match.end() + 100)
            ]
            attachment_match = re.search(
                r"(?:attachment|retention)[:\s]+\$?[\d,]+(?:\.\d+)?[MK]?",
                context,
                re.IGNORECASE,
            )
            attachment = 0
            if attachment_match:
                attachment = parse_currency(
                    attachment_match.group().split(":")[-1].strip()
                )

            limits.append(
                {
                    "limit": limit,
                    "attachment": attachment,
                    "is_primary": attachment == 0,
                    "raw_text": match_text.strip(),
                }
            )

    return limits


def extract_part_of_patterns(text: str) -> List[Dict[str, Any]]:
    """
    Extract "part of" patterns that show carrier-specific limit allocations.

    Examples:
    - "Ironshore Limits: $2,500,000 (being 3.333%) part of $75,000,000 excess of $100,000,000"
    - "Carrier A: $1M (25%) part of $4M xs $1M"
    - "Policy Limit: $5,000,000 that being 6.67% Annual Aggregate; part of $75,000,000 Excess of $100,000,000"

    Args:
        text: Text content to search

    Returns:
        List of dictionaries with carrier limits and layer information
    """
    results = []

    # Pattern 1: "that being X%" format (percentage NOT in parentheses)
    # Example: "Policy Limit: $5,000,000 that being 6.67% Annual Aggregate; part of $75,000,000 Excess of $100,000,000"
    that_being_pattern = r"([A-Za-z][A-Za-z\s&\'.]+?)(?:Limits?)?:\s*\$?([\d,]+(?:\.\d+)?[MK]?)\s+that\s+being\s+([\d.]+)%[^;]*?;\s*part\s+of\s+\$?([\d,]+(?:\.\d+)?[MK]?)\s*(?:excess\s+of|xs|x/s)\s*\$?([\d,]+(?:\.\d+)?[MK]?)"

    that_being_matches = re.finditer(that_being_pattern, text, re.IGNORECASE)

    for match in that_being_matches:
        carrier_name = match.group(1).strip()
        carrier_amount = parse_currency(match.group(2))
        percentage = float(match.group(3))
        layer_limit = parse_currency(match.group(4))
        attachment = parse_currency(match.group(5))

        results.append(
            {
                "carrier_name": carrier_name,
                "carrier_limit": carrier_amount,
                "share": percentage / 100.0,
                "layer_limit": layer_limit,
                "attachment": attachment,
                "is_primary": False,
                "raw_text": match.group(0).strip(),
            }
        )

    # Pattern 2: Original pattern with parentheses - "(being X%)" or "(X%)"
    # Example: "Ironshore Limits: $2,500,000 (being 3.333%) part of $75,000,000 excess of $100,000,000"
    part_of_pattern = r"([A-Za-z][A-Za-z\s&\'.]+?)(?:Limits?)?:\s*\$?([\d,]+(?:\.\d+)?[MK]?)\s*\((?:being\s+)?([\d.]+)%\)\s*part\s+of\s+\$?([\d,]+(?:\.\d+)?[MK]?)\s*(?:excess\s+of|xs|x/s)\s*\$?([\d,]+(?:\.\d+)?[MK]?)"

    matches = re.finditer(part_of_pattern, text, re.IGNORECASE)

    for match in matches:
        carrier_name = match.group(1).strip()
        carrier_amount = parse_currency(match.group(2))
        percentage = float(match.group(3))
        layer_limit = parse_currency(match.group(4))
        attachment = parse_currency(match.group(5))

        results.append(
            {
                "carrier_name": carrier_name,
                "carrier_limit": carrier_amount,
                "share": percentage / 100.0,
                "layer_limit": layer_limit,
                "attachment": attachment,
                "is_primary": False,
                "raw_text": match.group(0).strip(),
            }
        )

    # Pattern 3: Simple "that being" without carrier name
    # Example: "$5,000,000 that being 6.67% Annual Aggregate; part of $75,000,000 Excess of $100,000,000"
    simple_that_being = r"\$?([\d,]+(?:\.\d+)?[MK]?)\s+that\s+being\s+([\d.]+)%[^;]*?;\s*part\s+of\s+\$?([\d,]+(?:\.\d+)?[MK]?)\s*(?:excess\s+of|xs|x/s)\s*\$?([\d,]+(?:\.\d+)?[MK]?)"

    simple_that_matches = re.finditer(simple_that_being, text, re.IGNORECASE)

    for match in simple_that_matches:
        # Skip if already captured by the detailed pattern
        if any(r["raw_text"] in match.group(0) for r in results):
            continue

        carrier_amount = parse_currency(match.group(1))
        percentage = float(match.group(2))
        layer_limit = parse_currency(match.group(3))
        attachment = parse_currency(match.group(4))

        results.append(
            {
                "carrier_name": "Unknown Carrier",
                "carrier_limit": carrier_amount,
                "share": percentage / 100.0,
                "layer_limit": layer_limit,
                "attachment": attachment,
                "is_primary": False,
                "raw_text": match.group(0).strip(),
            }
        )

    # Pattern 4: Simple "part of" with parentheses, without carrier name
    # Example: "$X (Y%) part of $Z excess of $W"
    simple_part_of = r"\$?([\d,]+(?:\.\d+)?[MK]?)\s*\((?:being\s+)?([\d.]+)%\)\s*part\s+of\s+\$?([\d,]+(?:\.\d+)?[MK]?)\s*(?:excess\s+of|xs|x/s)\s*\$?([\d,]+(?:\.\d+)?[MK]?)"

    simple_matches = re.finditer(simple_part_of, text, re.IGNORECASE)

    for match in simple_matches:
        # Skip if already captured by the detailed pattern
        if any(r["raw_text"] in match.group(0) for r in results):
            continue

        carrier_amount = parse_currency(match.group(1))
        percentage = float(match.group(2))
        layer_limit = parse_currency(match.group(3))
        attachment = parse_currency(match.group(4))

        results.append(
            {
                "carrier_name": "Unknown Carrier",
                "carrier_limit": carrier_amount,
                "share": percentage / 100.0,
                "layer_limit": layer_limit,
                "attachment": attachment,
                "is_primary": False,
                "raw_text": match.group(0).strip(),
            }
        )

    # Pattern 5: Primary "that being" patterns
    # Example: "Carrier: $X that being Y% [text]; part of $Z Primary"
    primary_that_being = r"([A-Za-z][A-Za-z\s&\'.]+?)(?:Limits?)?:\s*\$?([\d,]+(?:\.\d+)?[MK]?)\s+that\s+being\s+([\d.]+)%[^;]*?;\s*part\s+of\s+\$?([\d,]+(?:\.\d+)?[MK]?)\s*(?:primary|primary\s+layer)"

    primary_that_matches = re.finditer(primary_that_being, text, re.IGNORECASE)

    for match in primary_that_matches:
        carrier_name = match.group(1).strip()
        carrier_amount = parse_currency(match.group(2))
        percentage = float(match.group(3))
        layer_limit = parse_currency(match.group(4))

        results.append(
            {
                "carrier_name": carrier_name,
                "carrier_limit": carrier_amount,
                "share": percentage / 100.0,
                "layer_limit": layer_limit,
                "attachment": 0,
                "is_primary": True,
                "raw_text": match.group(0).strip(),
            }
        )

    # Pattern 6: Primary "part of" with parentheses
    # Example: "Carrier: $X (Y%) part of $Z Primary"
    primary_part_of = r"([A-Za-z][A-Za-z\s&\'.]+?)(?:Limits?)?:\s*\$?([\d,]+(?:\.\d+)?[MK]?)\s*\((?:being\s+)?([\d.]+)%\)\s*part\s+of\s+\$?([\d,]+(?:\.\d+)?[MK]?)\s*(?:primary|primary\s+layer)"

    primary_matches = re.finditer(primary_part_of, text, re.IGNORECASE)

    for match in primary_matches:
        carrier_name = match.group(1).strip()
        carrier_amount = parse_currency(match.group(2))
        percentage = float(match.group(3))
        layer_limit = parse_currency(match.group(4))

        results.append(
            {
                "carrier_name": carrier_name,
                "carrier_limit": carrier_amount,
                "share": percentage / 100.0,
                "layer_limit": layer_limit,
                "attachment": 0,
                "is_primary": True,
                "raw_text": match.group(0).strip(),
            }
        )

    return results


def extract_carrier_info(text: str) -> List[Dict[str, Any]]:
    """
    Extract carrier information from text.
    Looks for carrier names, shares, and premiums.

    Args:
        text: Text content to search

    Returns:
        List of carrier information dictionaries
    """
    carriers = []

    # Common insurance carrier patterns
    carrier_keywords = [
        "insurance",
        "assurance",
        "indemnity",
        "underwriters",
        "syndicate",
        "lloyd",
        "mutual",
        "casualty",
        "risk",
    ]

    # Pattern: Carrier name followed by percentage or share
    # Example: "ABC Insurance Company - 50%"
    lines = text.split("\n")
    for i, line in enumerate(lines):
        # Look for percentage patterns
        percentage_match = re.search(r"(\d+(?:\.\d+)?)\s*%", line)

        if percentage_match:
            # Check if line contains carrier keywords
            line_lower = line.lower()
            has_carrier_keyword = any(
                keyword in line_lower for keyword in carrier_keywords
            )

            if has_carrier_keyword or len(line.split()) > 2:
                # Extract carrier name (text before the percentage)
                carrier_name_match = re.search(r"^(.+?)[\s-]*\d+(?:\.\d+)?\s*%", line)
                if carrier_name_match:
                    carrier_name = carrier_name_match.group(1).strip()
                    share = float(percentage_match.group(1)) / 100.0

                    # Look for premium in nearby lines
                    premium = 0
                    context_lines = lines[max(0, i - 2) : min(len(lines), i + 3)]
                    for context_line in context_lines:
                        premium_match = re.search(
                            r"(?:premium|prem)[:\s]+\$?[\d,]+(?:\.\d+)?",
                            context_line,
                            re.IGNORECASE,
                        )
                        if premium_match:
                            premium = parse_currency(premium_match.group())
                            break

                    carriers.append(
                        {
                            "carrier_name": carrier_name,
                            "share": share,
                            "premium": premium,
                            "raw_text": line.strip(),
                        }
                    )

    return carriers


def extract_policy_number(text: str) -> Optional[str]:
    """
    Extract policy number from text.

    Args:
        text: Text content to search

    Returns:
        Policy number if found, None otherwise
    """
    # Common policy number patterns
    patterns = [
        r"(?:policy\s+(?:no|number|#))[:\s]+([A-Z0-9-]+)",
        r"(?:certificate\s+(?:no|number|#))[:\s]+([A-Z0-9-]+)",
        r"(?:binder\s+(?:no|number|#))[:\s]+([A-Z0-9-]+)",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()

    return None


def parse_insurance_pdf(pdf_bytes: bytes, filename: str = "") -> Dict[str, Any]:
    """
    Parse an insurance PDF document and extract structured data.

    Args:
        pdf_bytes: PDF file as bytes
        filename: Original filename (optional, for reference)

    Returns:
        Dictionary containing extracted insurance data
    """
    try:
        # Extract text from PDF
        text = extract_text_from_pdf(pdf_bytes)

        # Extract "part of" patterns (carrier-specific limit allocations)
        # These are more specific and take priority
        part_of_data = extract_part_of_patterns(text)

        # Extract limits and layers
        limits = extract_limit_patterns(text)

        # Extract carrier information
        carriers = extract_carrier_info(text)

        # Extract policy number
        policy_number = extract_policy_number(text)

        # Determine document type
        doc_type = "Unknown"
        text_lower = text.lower()
        if "quote" in text_lower:
            doc_type = "Quote"
        elif "binder" in text_lower:
            doc_type = "Binder"
        elif "policy" in text_lower:
            doc_type = "Policy"
        elif "certificate" in text_lower:
            doc_type = "Certificate"

        return {
            "filename": filename,
            "document_type": doc_type,
            "policy_number": policy_number,
            "limits": limits,
            "carriers": carriers,
            "part_of_data": part_of_data,  # Carrier-specific limit allocations
            "raw_text": text[:1000],  # First 1000 chars for preview
            "success": True,
            "error": None,
        }

    except Exception as e:
        return {
            "filename": filename,
            "success": False,
            "error": str(e),
            "limits": [],
            "carriers": [],
        }


def merge_parsed_documents(parsed_docs: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Merge multiple parsed documents into a single program structure.

    Args:
        parsed_docs: List of parsed document dictionaries

    Returns:
        Program structure with merged layers and carriers
    """
    # Collect all unique limits/layers
    layer_dict = {}  # Key: (limit, attachment) -> layer data

    for doc in parsed_docs:
        if not doc.get("success"):
            continue

        # PRIORITY 1: Process "part of" patterns (most specific)
        # These contain both layer and carrier info together
        for part_of in doc.get("part_of_data", []):
            # Use the layer_limit (total limit) not carrier_limit
            key = (part_of["layer_limit"], part_of["attachment"])

            if key not in layer_dict:
                layer_dict[key] = {
                    "limit": part_of["layer_limit"],
                    "attachment": part_of["attachment"],
                    "is_primary": part_of["is_primary"],
                    "carriers": [],
                }

            # Check if carrier already exists in this layer
            existing_carrier = None
            for existing in layer_dict[key]["carriers"]:
                if existing["carrier_name"] == part_of["carrier_name"]:
                    existing_carrier = existing
                    break

            if existing_carrier:
                # Update existing carrier
                existing_carrier["premium"] += part_of.get("carrier_limit", 0)
            else:
                # Add new carrier with their specific share
                layer_dict[key]["carriers"].append(
                    {
                        "carrier_name": part_of["carrier_name"],
                        "share": part_of["share"],
                        "premium": part_of.get(
                            "carrier_limit", 0
                        ),  # Use carrier limit as premium estimate
                        "carrier_fee": 0.0,
                        "surplus_fee": 0.0,
                        "policy_number": doc.get("policy_number", ""),
                        "has_multiple_rbes": False,
                        "rbes": [],
                    }
                )

        # PRIORITY 2: Process regular limit patterns
        for limit_info in doc.get("limits", []):
            key = (limit_info["limit"], limit_info["attachment"])

            if key not in layer_dict:
                layer_dict[key] = {
                    "limit": limit_info["limit"],
                    "attachment": limit_info["attachment"],
                    "is_primary": limit_info["is_primary"],
                    "carriers": [],
                }

            # Add carriers from this document to this layer
            # Only if we don't already have carriers from "part of" data
            if not layer_dict[key]["carriers"]:
                for carrier in doc.get("carriers", []):
                    # Check if carrier already exists
                    existing_carrier = None
                    for existing in layer_dict[key]["carriers"]:
                        if existing["carrier_name"] == carrier["carrier_name"]:
                            existing_carrier = existing
                            break

                    if existing_carrier:
                        # Update existing carrier (sum premiums)
                        existing_carrier["premium"] += carrier.get("premium", 0)
                    else:
                        # Add new carrier
                        layer_dict[key]["carriers"].append(
                            {
                                "carrier_name": carrier.get("carrier_name", "Unknown"),
                                "share": carrier.get("share", 0),
                                "premium": carrier.get("premium", 0),
                                "carrier_fee": 0.0,
                                "surplus_fee": 0.0,
                                "policy_number": doc.get("policy_number", ""),
                                "has_multiple_rbes": False,
                                "rbes": [],
                            }
                        )

    # Convert to list and sort by attachment
    layers = sorted(layer_dict.values(), key=lambda x: x["attachment"])

    return {
        "layers": layers,
        "documents_processed": len([d for d in parsed_docs if d.get("success")]),
        "documents_failed": len([d for d in parsed_docs if not d.get("success")]),
    }
