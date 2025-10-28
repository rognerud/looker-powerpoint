import json
import hashlib
from collections.abc import Mapping


def _to_plain(query):
    # Convert common model types to plain dicts for stable serialization
    if hasattr(query, "model_dump"):
        return query.model_dump(exclude_none=True)
    if hasattr(query, "dict"):
        return query.dict(exclude_none=True)
    # If it's already a mapping, cast to dict to remove custom Mapping types
    if isinstance(query, Mapping):
        return dict(query)
    return query


def _hash_query_obj(obj) -> str:
    # Create a canonical JSON string and hash it
    canonical = json.dumps(obj, sort_keys=True, separators=(",", ":"), default=str)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def group_queries_by_identity(queries):
    """
    Groups identical queries so they can be executed once.
    Accepts 'queries' as a mapping of shape_id -> query (so we iterate queries.items()).
    Returns a list with the same shape: {"query": original_query, "shapes": [...]}
    """
    groups = {}
    for shape_id, query in queries.items():
        # Keep the original query object for downstream use
        original_query = query

        # Convert model objects to plain serializable form for hashing only
        query_obj = _to_plain(query)

        # Compute stable canonical hash
        key = _hash_query_obj(query_obj)

        if key not in groups:
            groups[key] = {"query": original_query, "shapes": []}
        groups[key]["shapes"].append(shape_id)

    return list(groups.values())
