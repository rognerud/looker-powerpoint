from deepdiff import DeepHash

def group_queries_by_identity(queries):
    """
    Groups identical queries so they can be executed once.
    Works for dicts, Pydantic objects, and Looker SDK WriteQuery models.
    """
    groups = {}

    for shape_id, query in queries.items():
        # Flatten model objects to dicts if possible
        if hasattr(query, "model_dump"):
            query_obj = query.model_dump(exclude_none=True)
        elif hasattr(query, "dict"):
            query_obj = query.dict(exclude_none=True)
        else:
            query_obj = query

        # Compute a stable deep hash (no extra args needed)
        key = DeepHash(query_obj)[query_obj]

        if key not in groups:
            groups[key] = {"query": query, "shapes": []}
        groups[key]["shapes"].append(shape_id)

    return list(groups.values())