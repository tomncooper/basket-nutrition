import requests

import pandas as pd

PRODUCT_SEARCH_URL: str = "https://dev.tescolabs.com/grocery/products"

PRODUCT_DATA_URL: str = "https://dev.tescolabs.com/product"


def product_search(api_key: str, query: str, limit: int = 10, offset: int = 0):

    headers = {"Ocp-Apim-Subscription-Key": api_key}

    params = {"query": query, "limit": limit, "offset": offset}

    response: requests.Response = requests.get(
        PRODUCT_SEARCH_URL, headers=headers, params=params
    )

    return response.json()


def product_data(api_key: str, product_id: str):

    headers = {"Ocp-Apim-Subscription-Key": api_key}

    params = {"tpnc": product_id}

    response: requests.Response = requests.get(
        PRODUCT_DATA_URL, headers=headers, params=params
    )

    return response.json()
