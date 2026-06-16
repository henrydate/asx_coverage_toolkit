"""
ASX index constituent sets — March 2026 rebalance.
Used for index membership flags in the enrichment pipeline.
"""

ASX_20: set[str] = {
    "CBA", "BHP", "WBC", "NAB", "ANZ", "WES", "CSL", "MQG", "GMG", "TLS",
    "RIO", "WDS", "TCL", "NST", "WOW", "FMG", "ALL", "BXB", "EVN", "QBE",
}

ASX_50: set[str] = ASX_20 | {
    "COL", "SCG", "RMD", "STO", "ORG", "COH", "REA", "S32", "AGL", "APA",
    "JHX", "DXS", "CPU", "AMC", "ASX", "SHL", "SUN", "IAG", "MGR", "XRO",
    "SGP", "GPT", "MIN", "TWE", "ALX", "QAN", "AZJ", "BLD", "CHC", "SOL",
}

ASX_100: set[str] = ASX_50 | {
    "VCX", "NEC", "MPL", "NHF", "BOQ", "BEN", "SVW", "WOR", "SEK", "WTC",
    "CAR", "ILU", "ORA", "ANN", "VEA", "IGO", "BSL", "IEL", "JBH", "HVN",
    "NWL", "PPT", "TNE", "NXT", "CQR", "IPL", "CWY", "DOW", "IPH", "PME",
    "ALQ", "FLT", "GNC", "MTS", "HLS", "CTD", "RRL", "SUL", "NHC", "LYC",
    "BAP", "BKL", "CCP", "SDF", "AWC", "PDL", "BKW", "PMV", "SGM", "APE",
}

ASX_200: set[str] = ASX_100 | {
    "WHC", "DHG", "BRG", "NSR", "ING", "GOR", "ARB", "INA", "CLW", "CMW",
    "RWC", "NUF", "RHC", "FPH", "MFG", "BPT", "QUB", "GOZ", "IFL", "SBM",
    "TPG", "ALD", "GEM", "IFT", "PLS", "LTR", "A2M", "ZIP", "MP1", "PBH",
    "360", "GQG", "HUB", "KLS", "DRR", "GDI", "ABB", "NST", "SFR", "YAL",
    "EMR", "BWP", "LOV", "KAR", "CMM", "CLW", "HDN", "GOZ", "HMC", "CIA",
    "DTL", "CDA", "LNW", "GYG", "EDV", "EBO", "BOE", "AUB", "JDO", "CHN",
    "NAN", "CHC", "CQR", "CRN", "ASB", "EHL", "MMA",
}

ALL_ORDINARIES: set[str] = ASX_200 | {
    "ACL", "ADH", "AEF", "AGI", "ALC", "ALI", "ATA", "ATL", "AUB", "AUI",
    "AX1", "AWF", "BBN", "BCI", "BGP", "BIM", "BIS", "BKI", "BOL", "BRL",
    "BRN", "BSA", "BTH", "BUB", "BWX", "CAA", "CAM", "CBS", "CCX", "CDM",
    "CEL", "CEN", "CGC", "CGF", "CIN", "CIP", "CLQ", "CML", "CNN", "CNU",
    "COE", "COG", "COM", "COR", "CRD", "CRL", "DDR", "DGL", "DMP", "DOC",
    "DRE", "DSK", "DUG", "DYL", "EFG", "EGL", "ELO", "EML", "ENR", "EWC",
    "EXL", "FAL", "FBR", "FCT", "FDV", "FLG", "FMG", "FND", "FRI", "FSA",
    "GAL", "GDF", "GFF", "GLN", "GNG", "GOW", "GRR", "GS1", "GUD", "HAV",
    "HCW", "HDN", "HGO", "HIT", "HLS", "HOM", "HVN", "IKE", "ILU", "IMA",
    "IMD", "IMM", "IND", "INR", "ION", "IPD", "IPC", "IRE", "IVZ", "JBH",
    "JHX", "JIN", "JMS", "JNS", "KAR", "KGN", "KLS", "KMD", "KNO", "KSC",
    "LAU", "LBL", "LDR", "LEL", "LEX", "LFG", "LGI", "LIC", "LIN", "LIT",
    "LLC", "LMG", "LOT", "LOV", "LPE", "LRK", "LTR", "LYL", "MAD", "MAF",
    "MAH", "MAQ", "MBH", "MEI", "MEL", "MFD", "MGH", "MGL", "MLS", "MLX",
    "MMI", "MND", "MPL", "MQR", "MSB", "MTC", "MTM", "MTS", "MYE", "NAN",
    "NEC", "NEA", "NHC", "NHF", "NSR", "NUF", "NWH", "NWL", "NXT", "OFX",
    "ORA", "ORG", "ORI", "OZL", "PDL", "PLS", "PME", "PMV", "PNV", "PPT",
    "PTM", "QAN", "QBE", "QUB", "REA", "RMD", "RRL", "RWC", "SBM", "SCG",
    "SCP", "SDR", "SEK", "SFR", "SGM", "SGP", "SHL", "SIG", "SLR", "SOL",
    "SSR", "SUL", "SUN", "SVW", "TAH", "TGR", "TLC", "TLS", "TPG", "TWE",
    "UNI", "VCX", "VEA", "VOC", "VVR", "WAF", "WBC", "WEB", "WES", "WHC",
    "WOR", "WOW", "WTC", "XRO", "YAL", "ZIP",
}


def get_index_flags(ticker: str) -> dict[str, str]:
    """Return Yes/No flags and concatenated membership string for a ticker."""
    in_20  = ticker in ASX_20
    in_50  = ticker in ASX_50
    in_100 = ticker in ASX_100
    in_200 = ticker in ASX_200
    in_ao  = ticker in ALL_ORDINARIES

    parts = []
    if in_20:  parts.append("ASX 20")
    if in_50:  parts.append("ASX 50")
    if in_100: parts.append("ASX 100")
    if in_200: parts.append("ASX 200")
    if in_ao:  parts.append("All Ords")

    return {
        "ASX 20":        "Yes" if in_20  else "No",
        "ASX 50":        "Yes" if in_50  else "No",
        "ASX 100":       "Yes" if in_100 else "No",
        "ASX 200":       "Yes" if in_200 else "No",
        "All Ordinaries":"Yes" if in_ao  else "No",
        "Index Membership": " | ".join(parts) if parts else "Outside Major Indices",
    }
