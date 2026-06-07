"""
SAP Object Configuration — V4
Defines join keys and key fields for each SAP migration object.
Add new objects at the bottom — no other changes needed.
"""

from pathlib import Path

SAP_OBJECT_CONFIG = {
    "MATERIAL": {
        "join_key":    "MATNR",
        "description": "Material Master",
        "key_fields":  ["MATNR","MAKTX","MTART","MATKL","MEINS",
                        "BRGEW","NTGEW","STPRS","VPRSV","WERKS"],
    },
    "VENDOR": {
        "join_key":    "LIFNR",
        "description": "Vendor Master",
        "key_fields":  ["LIFNR","NAME1","KTOKK","LAND1","STRAS",
                        "ORT01","PSTLZ","ZTERM","AKONT","WAERS"],
    },
    "CUSTOMER": {
        "join_key":    "KUNNR",
        "description": "Customer Master",
        "key_fields":  ["KUNNR","NAME1","KTOKD","LAND1","STRAS",
                        "ORT01","PSTLZ","ZTERM","WAERS","VKORG"],
    },
    "BUSINESS_PARTNER": {
        "join_key":    "KUNNR",
        "description": "Business Partner",
        "key_fields":  ["KUNNR","NAME1","NAME2","KTOKD","LAND1","STRAS",
                        "ORT01","PSTLZ","REGIO","ZTERM","WAERS","SPRAS"],
    },
    "GL_ACCOUNT": {
        "join_key":    "SAKNR",
        "description": "GL Account / Chart of Accounts",
        "key_fields":  ["SAKNR","BUKRS","TXT20","TXT50","KTOKS","XBILK","FSTAG"],
    },
    "OPEN_ITEMS_AR": {
        "join_key":    "BELNR",
        "description": "Open Items — Accounts Receivable",
        "key_fields":  ["BELNR","GJAHR","BUZEI","KUNNR","BLDAT",
                        "BUDAT","WRBTR","DMBTR","ZTERM","ZFBDT"],
    },
    "OPEN_ITEMS_AP": {
        "join_key":    "BELNR",
        "description": "Open Items — Accounts Payable",
        "key_fields":  ["BELNR","GJAHR","BUZEI","LIFNR","BLDAT",
                        "BUDAT","WRBTR","DMBTR","ZTERM","ZFBDT"],
    },
    "PURCHASE_ORDER": {
        "join_key":    "EBELN",
        "description": "Purchase Orders",
        "key_fields":  ["EBELN","EBELP","LIFNR","MATNR","MENGE",
                        "MEINS","NETPR","PEINH","WAERS","WERKS"],
    },
    "SALES_ORDER": {
        "join_key":    "VBELN",
        "description": "Sales Orders",
        "key_fields":  ["VBELN","POSNR","KUNNR","MATNR","KWMENG",
                        "VRKME","NETWR","WAERS","WERKS","VKORG"],
    },
    "ASSET": {
        "join_key":    "ANLN1",
        "description": "Fixed Asset Master",
        "key_fields":  ["ANLN1","ANLN2","BUKRS","ANLKL","TXT50",
                        "AKTIV","DEAKT","KOSTL","AUFNR","WAERS"],
    },
    "COST_CENTRE": {
        "join_key":    "KOSTL",
        "description": "Cost Centre Master",
        "key_fields":  ["KOSTL","BUKRS","KOKRS","KTEXT","KOSAR",
                        "ABTEI","VERAK","WAERS","DATAB","DATBI"],
    },
    "PROFIT_CENTRE": {
        "join_key":    "PRCTR",
        "description": "Profit Centre Master",
        "key_fields":  ["PRCTR","KOKRS","KTEXT","LTEXT","ABTEI",
                        "VERAK","WAERS","DATAB","DATBI"],
    },
    "BANK": {
        "join_key":    "BANKL",
        "description": "Bank Master",
        "key_fields":  ["BANKL","BANKS","BANKA","STRAS","ORT01","SWIFT","BGRUP"],
    },
    "INVENTORY": {
        "join_key":    "MATNR",
        "description": "Inventory / Stock Balances",
        "key_fields":  ["MATNR","WERKS","LGORT","LABST","INSME",
                        "EINME","SPEME","MEINS","STPRS","WAERS"],
    },
}


def get_object_config(name: str) -> dict:
    """
    Auto-detect SAP object config from a table/file name.
    Tries exact match first, then longest partial match.

    Examples:
        MATERIAL              -> MATERIAL config
        SOURCE_CUSTOMERS      -> CUSTOMER config
        VENDOR_DATA           -> VENDOR config
        customer_master       -> CUSTOMER config
    """
    stem = str(name).upper().replace("-", "_").replace(" ", "_")

    if stem in SAP_OBJECT_CONFIG:
        return SAP_OBJECT_CONFIG[stem]

    matches = [(k, v) for k, v in SAP_OBJECT_CONFIG.items() if k in stem]
    if matches:
        best = sorted(matches, key=lambda x: len(x[0]), reverse=True)[0]
        return best[1]

    return {}
