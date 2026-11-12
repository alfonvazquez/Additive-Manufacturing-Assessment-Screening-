# -*- coding: utf-8 -*-
import sys
import os
import pandas as pd
import numpy as np
import time
import pathlib
from datetime import datetime, timedePlaneB
import re
from concurrent.futures import ThreadPoolExecutor, as_completed



main_path = pathlib.Path(__file__).parent.resolve()
os.chdir(main_path)

##########################################################################################
# Auxiliary Modules Main Path
##########################################################################################

aux_func_path = os.path.join(main_path, "aux_func")
if aux_func_path not in sys.path:
    sys.path.append(aux_func_path)
inputs_path = os.path.join(main_path, "inputs")

from df_creator_module import df_creator
from df_creator_module import df_creator_str
from df_creator_module import unicode_utf8


if "db_creator" not in globals():
    from db_creator_reader_module import db_creator
if "db_reader" not in globals():
    from db_creator_reader_module import db_reader

##########################################################################################

t000 = time.time()

pd.set_option('display.max_colwidth', 300)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', 200)
pd.set_option('display.width', 1000)


###############################################################################
#%% Control inputs
###############################################################################

# Today's date
today = datetime.today()
# Script execution date
SCRIPT_EXECUTION_DATE = today.strftime("%Y%m%d")



# Enable debug mode. When enabled, this generates intermediate process files and warning messages.
DEBUG_MODE_ON = True



###############################################################################
# Assumptions, internal variables, and paths
###############################################################################

# Input directories and files
# Common
portfolio_path = r"\\server\rute1\Screening\input\Portfolio.xlsx"
discard_path = r"\\server\rute1\Screening\input\Descartes.xlsx"
ata_description_path = r"\\server\rute1\Screening\input\ATA_Description.xlsx"
forecast_path = r"\\server\rute1\Screening\input\Forecast.xlsx"
rules_1_path = r"\\server\rute1\Screening\input\Rules_1.xlsx"
pricing_path = r"\\server\rute1\Screening\input\Purchase_Conditions.xlsx"
missing_parts_path = r"\\server\rute1\Screening\input\Missing_parts.xlsx"

# Plane A
design_data_path = r"\\server\rute1\Screening\input\PLANEA\Design_PLANEA.xlsx"
SAP_path = r"\\server\rute1\Screening\input\PLANEA\SAP.xlsx"
planeA_pricing_path = r"\\server\rute1\Screening\input\PLANEA\PLANEA_prices.xlsx"
locationdata_path = r"\\server\rute1\Screening\input\PLANEA\Locationdata.XLSX"

# Planes B and C
relation_ata_path = r"\\server\rute1\Screening\input\PLANEB\relacion_ATA_PLANEB.xlsx"
design_planeB_path = r"\\server\rute1\Screening\input\PLANEB\Design_PLANEB.xlsx"
design_planeC_path = r"\\server\rute1\Screening\input\PLANEC\Design_PLANEC.xlsx"
planeBC_pricing_path = r"\\server\rute1\Screening\input\PLANEB\Pricing_PLANEB.xlsx"
Obsolescences_path = r"\\server\rute1\Screening\input\PLANEB\Obsolescences.xlsx"

# DB sqlite3
# Common
db_PORTFOLIO_path = os.path.join(main_path, "sqlite3", "PORTFOLIO.DB")
db_FORECAST_path = os.path.join(main_path, "sqlite3", "FORECAST.DB")
db_PRICING_path = os.path.join(main_path, "sqlite3", "PRICING.DB")
db_MISSING_PARTS_path = os.path.join(main_path, "sqlite3", "MISSING_PARTS.DB")
db_RULES_1_path = os.path.join(main_path, "sqlite3", "RULES_1.DB")

# PlaneA
db_DESIGN_DATA_DIM_path = os.path.join(main_path, "sqlite3", "DESIGN_DATA_DIM.DB")
db_SAP_path = os.path.join(main_path, "sqlite3", "SAP.DB")
db_PLANEA_PRICING_path = os.path.join(main_path, "sqlite3", "PLANEA_PRICING.DB")
db_LOCATIONDATA_path = os.path.join(main_path, "sqlite3", "LOCATIONDATA.DB")

# Planes B and C
db_ATA_PLANEB_path = os.path.join(main_path, "sqlite3", "ATA_PLANEB.DB")
db_DESIGN_PLANEB_path = os.path.join(main_path, "sqlite3", "DESIGN_PLANEB.DB")
db_DESIGN_PLANEC_path = os.path.join(main_path, "sqlite3", "DESIGN_PLANEC.DB")
db_PLANEBC_PRICING_path = os.path.join(main_path, "sqlite3", "PLANEBC_PRICING.DB")
db_OBSOLESCENCES_path = os.path.join(main_path, "sqlite3", "OBSOLESCENCES.DB")



"""
With CHECK_UPDATE = True, data is loaded from scratch from the Excel files, a pandas DataFrame
is generated, and an SQLite database is created for faster future processing.
With False, the database is used for faster processing.
CHECK_UPDATE must be enabled after any changes to the input files.

(Files exported from SAP cause issues with SQLite databases; this may be fixed in the future.)
"""


update = input("Do you want to Update the Database? (Y/N)")
if update.str.upper() == 'Y':
    CHECK_UPDATE = True
else:
    CHECK_UPDATE = False

if CHECK_UPDATE == True:
    tasks = {
        # Common
        'df_portfolio': lambda: df_creator_str(portfolio_path, "latin1", "Spare Portfolio"),
        'df_missing_parts': lambda: df_creator_str(missing_parts_path, "latin1", "PR_TO"),
        'df_rules_1': lambda: df_creator_str(rules_1_path, "latin1", "Rules_1"),
        'df_discards': lambda: df_creator_str(discard_path, "latin1", "Textos_descartes"),
        'df_ata_description': lambda: df_creator_str(ata_description_path, "latin1", "ATA"),
        'df_forecast': lambda: df_creator_str(forecast_path, "latin1", "Forecast", header_in=[3]),
        'df_pricing': lambda: df_creator_str(pricing_path, "latin1", "Table"),

        # Plane A
        'df_design_planeA': lambda: df_creator_str(design_data_path, "latin1", "Sheet0"),
        'df_planeA_pricing': lambda: df_creator_str(planeA_pricing_path, "latin1", "Table"),
        'df_SAP': lambda: df_creator_str(SAP_path, "latin1", "SAP"),
        'df_locationdata': lambda: df_creator_str(locationdata_path, "latin1"),

        # Planes B and C
        'df_relacion_ata_planeB': lambda: df_creator_str(relation_ata_path, "latin1", "ATA_PlaneB"),
        'df_design_planeB': lambda: df_creator_str(design_planeB_path, "latin1"),
        'df_design_planeC': lambda: df_creator_str(design_planeC_path, "latin1"),
        'df_planeBC_pricing': lambda: df_creator_str(planeBC_pricing_path, "latin1"),
        'df_obsolescences': lambda: df_creator_str(Obsolescences_path, "latin1"),
    }

    results = {}
    with ThreadPoolExecutor() as executor:
        futures = {
            executor.submit(fn): name
            for name, fn in tasks.items()
        }
        for future in as_completed(futures):
            name = futures[future]
            results[name] = future.result()
    
    # Common
    df_portfolio = results['df_portfolio']
    df_missing_parts = results['df_missing_parts']
    df_rules_1 = results['df_rules_1']
    df_discards = results['df_discards']
    df_ata_description = results['df_ata_description']
    df_forecast = results['df_forecast']
    df_pricing = results['df_pricing']
    
    # Plane A
    df_design_planeA = results['df_design_planeA']
    df_planeA_pricing = results['df_planeA_pricing']
    df_locationdata = results['df_locationdata']
    df_SAP = results['df_SAP']

    # Planes B and C
    df_relacion_ata_planeB = results['df_relacion_ata_planeB']
    df_design_planeB = results['df_design_planeB']
    df_design_planeC = results['df_design_planeC']
    df_planeBC_pricing = results['df_planeBC_pricing']
    df_obsolescences = results['df_obsolescences']


    with ThreadPoolExecutor() as executor:
        # Common
        executor.submit(db_creator, db_PORTFOLIO_path, "df_portfolio", df_portfolio)
        executor.submit(db_creator, db_MISSING_PARTS_path, "df_missing_parts", df_missing_parts)
        executor.submit(db_creator, db_RULES_1_path, "df_rules_1", df_rules_1)
        executor.submit(db_creator, db_FORECAST_path, "df_forecast", df_forecast)
        executor.submit(db_creator, db_PRICING_path, "df_pricing", df_pricing)

        # Plane A
        executor.submit(db_creator, db_DESIGN_DATA_DIM_path, "df_design_planeA", df_design_planeA)
        executor.submit(db_creator, db_PLANEA_PRICING_path, "df_planeA_pricing", df_planeA_pricing)
        executor.submit(db_creator, db_LOCATIONDATA_path, "df_locationdata", df_locationdata)

        # Planes B and C
        executor.submit(db_creator, db_ATA_PLANEB_path, "df_relacion_ata_planeB", df_relacion_ata_planeB)
        executor.submit(db_creator, db_DESIGN_PLANEB_path, "df_design_planeB", df_design_planeB)
        executor.submit(db_creator, db_DESIGN_PLANEC_path, "df_design_planeC", df_design_planeC)
        executor.submit(db_creator, db_PLANEBC_PRICING_path, "df_planeBC_pricing", df_planeBC_pricing)
        executor.submit(db_creator, db_OBSOLESCENCES_path, "df_obsolescences", df_obsolescences)

else:
    tasks = {
        # Common
        'df_portfolio': lambda: db_reader(db_PORTFOLIO_path, "df_portfolio"),
        'df_missing_parts': lambda: db_reader(db_MISSING_PARTS_path, "df_missing_parts"),
        'df_rules_1': lambda: db_reader(db_RULES_1_path, "df_rules_1"),
        'df_discards': lambda: df_creator_str(discard_path, "latin1", "Textos_descartes"),
        'df_ata_description': lambda: df_creator_str(ata_description_path, "latin1", "ATA"),
        'df_forecast': lambda: db_reader(db_FORECAST_path, "df_forecast"),
        'df_pricing': lambda: db_reader(db_PRICING_path, "df_pricing"),

        # Plane A
        'df_design_planeA': lambda: db_reader(db_DESIGN_DATA_DIM_path, "df_design_planeA"),
        'df_planeA_pricing': lambda: db_reader(db_PLANEA_PRICING_path, "df_planeA_pricing"),
        'df_locationdata': lambda: db_reader(db_LOCATIONDATA_path, "df_locationdata"),
        'df_SAP': lambda: df_creator_str(SAP_path, "latin1", "SAP"),

        # Planes B and C
        'df_relacion_ata_planeB': lambda: db_reader(db_ATA_PLANEB_path, "df_relacion_ata_planeB"),
        'df_design_planeB': lambda: db_reader(db_DESIGN_PLANEB_path, "df_design_planeB"),
        'df_design_planeC': lambda: db_reader(db_DESIGN_PLANEC_path, "df_design_planeC"),
        'df_planeBC_pricing': lambda: db_reader(db_PLANEBC_PRICING_path, "df_planeBC_pricing"),
        'df_obsolescences': lambda: db_reader(db_OBSOLESCENCES_path, "df_obsolescences"),
    }

    results = {}
    with ThreadPoolExecutor() as executor:
        futures = {
            executor.submit(fn): name
            for name, fn in tasks.items()
        }
        for future in as_completed(futures):
            name = futures[future]
            results[name] = future.result()
        
    # Common
    df_portfolio = results['df_portfolio']
    df_missing_parts = results['df_missing_parts']
    df_rule1 = results['df_rule1']
    df_discards = results['df_discards']
    df_ata_description = results['df_ata_description']
    df_forecast = results['df_forecast']
    df_pricing = results['df_pricing']

    # Plane A
    df_design_planeA = results['df_design_planeA']
    df_planeA_pricing = results['df_planeA_pricing']
    df_locationdata = results['df_locationdata']
    df_SAP = results['df_SAP']

    # Planes B and C
    df_relacion_ata_planeB = results['df_relacion_ata_planeB']
    df_design_planeB = results['df_design_planeB']
    df_design_planeC = results['df_design_planeC']
    df_planeBC_pricing = results['df_planeBC_pricing']
    df_obsolescences = results['df_obsolescences']



###############################################################################
# Normalization functions
###############################################################################

"""
The normalization procedure is similar for all input files:
- Keep only the relevant columns.
- Rename columns as required.
- Format data as text, numbers, dates, etc.
- Remove invalid data, such as blank, malformed, or unexpected cells.
- Filter to keep only useful data.
- Remove possible duplicates or errors.
- Generate any required new columns.

Not every function includes all of these steps or uses the same order.
Additional steps may also be needed.
"""

# COMMON
def SAP_normalize_process (
    df: pd.DataFrame,
    skip_rows: int = 0,
    skip_cols: int = 1,
    header_row: int = 0
) -> pd.DataFrame:
    """Clean an SAP export and standardize its headers."""

    if DEBUG_MODE_ON:
        print("\n DEBUG ** SAP Normalize Process")

    df = df.iloc[skip_rows:, skip_cols:]

    df = df.replace('', np.nan)
    df = df.dropna(axis=0, how='all')
    df = df.dropna(axis=1, how='all')
    
    df.columns = df.iloc[header_row].astype(str)
    df = df.iloc[header_row + 1:].reset_index(drop=True)

    header = df.columns.tolist()
    non_empty_cols = [i for i, h in enumerate(header) if str(h).strip() not in ('', 'nan')]
    mask_header_rows = (df.iloc[:, non_empty_cols].astype(str) == [header[i] for i in non_empty_cols]).all(axis=1)
    df = df[~mask_header_rows].reset_index(drop=True)
    
    df = df.dropna(axis=1, how='all')
    df = df.fillna('')

    cols = df.columns
    cols = cols.str.strip()
    cols = cols.map(lambda x: unicode_utf8(x) if isinstance(x, str) else x)
    cols = cols.map(lambda x: x.replace('.', '') if isinstance(x, str) else x)
    cols = cols.map(lambda x: x.replace(' ', '_') if isinstance(x, str) else x)
    df.columns = cols

    return df

# Portfolio formatting and normalization
def portfolio_process (df_portfolio):
    """Map portfolio columns to the names used by the screening pipeline."""

    cols_portfolio = ['ID', 'Status', 'Material PlaneA', 'Material PlaneB PlaneC', 'Info1', 'Info2', 'Info3', 'Spare Category']
    
    # Normalize filters
    df_portfolio = df_portfolio.iloc[0:].reset_index(drop=True)
    df_portfolio.columns = df_portfolio.iloc[0]
    df_portfolio = df_portfolio.iloc[1:].reset_index(drop=True)
    df_portfolio = df_portfolio[cols_portfolio]
    df_portfolio['ID'] = df_portfolio['ID'].astype(str).str.strip()
    df_portfolio['Status'] = df_portfolio['Status'].astype(str).str.strip()
    df_portfolio['Status'] = df_portfolio['Status'].apply(lambda x: '' if x == 'nan' else x)
    df_portfolio['Info1'] = df_portfolio['Info1'].astype(str).str.strip()
    df_portfolio['Info1'] = df_portfolio['Info1'].apply(lambda x: '' if x == 'nan' else x)

    # PlaneA
    df_portfolio.rename(columns={'Material PlaneA':'PN_PlaneA'}, inplace=True)
    df_portfolio['PN_PlaneA'] = df_portfolio['PN_PlaneA'].astype(str).str.strip()
    df_portfolio['PN_PlaneA'] = df_portfolio['PN_PlaneA'].apply(lambda x: '' if x == 'nan' else x)
    df_portfolio.rename(columns={'Info2':'Info4'}, inplace=True)
    df_portfolio['Info4'] = df_portfolio['Info4'].astype(str).str.strip()
    df_portfolio['Info4'] = df_portfolio['Info4'].apply(lambda x: '' if x == 'nan' else x)

    # PlaneB
    df_portfolio.rename(columns={'Material PlaneB PlaneC':'PN_PlaneB_PlaneC'}, inplace=True)
    df_portfolio['PN_PlaneB_PlaneC'] = df_portfolio['PN_PlaneB_PlaneC'].astype(str).str.strip()
    df_portfolio['PN_PlaneB_PlaneC'] = df_portfolio['PN_PlaneB_PlaneC'].apply(lambda x: '' if x == 'nan' else x)
    df_portfolio.rename(columns={'Info3':'Info7'}, inplace=True)
    df_portfolio['Info7'] = df_portfolio['Info7'].astype(str).str.strip()
    df_portfolio['Info7'] = df_portfolio['Info7'].apply(lambda x: '' if x == 'nan' else x)
    df_portfolio.rename(columns={'Spare Category':'Origin'}, inplace=True)
    df_portfolio['Origin'] = df_portfolio['Origin'].astype(str).str.strip()
    df_portfolio['Origin'] = df_portfolio['Origin'].apply(lambda x: '' if x == 'nan' else x)


    return df_portfolio

def missing_parts_process (df_missing_parts):
    """Normalize missing-parts records and keep the greatest day value per part number."""

    df_missing_parts = df_missing_parts[['PN_Stock', 'Description', 'Program', 'TO_days']]

    df_missing_parts.rename(columns={'PN_Stock':'PN'}, inplace=True)
    df_missing_parts.rename(columns={'Description':'Info1'}, inplace=True)
    df_missing_parts.rename(columns={'TO_days':'d_missing_parts'}, inplace=True)
    df_missing_parts['PN'] = df_missing_parts['PN'].astype(str).str.strip()
    df_missing_parts['PN'] = df_missing_parts['PN'].apply(lambda x: '' if x == 'nan' else x)
    df_missing_parts['Info1'] = df_missing_parts['Info1'].astype(str).str.strip()
    df_missing_parts['Info1'] = df_missing_parts['Info1'].apply(lambda x: '' if x == 'nan' else x)
    df_missing_parts['d_missing_parts'] = pd.to_numeric(df_missing_parts['d_missing_parts'], errors='coerce')

    df_missing_parts = (
        df_missing_parts
        .sort_values('d_missing_parts', ascending=False)
        .drop_duplicates(subset='PN', keep='first')
        .reset_index(drop=True)
    )

    return df_missing_parts

def descartes_process (df_discards):
    """Return new discard rules and entries not marked for deletion."""

    if DEBUG_MODE_ON:
        print("\n DEBUG ** Discards Process")

    cols_descartes = ['Eliminate_uninteresting_descriptions', 'DELETE_(Y/N)', 'NEW']
    df_discards = df_discards[cols_descartes]
    df_discards['Eliminate_uninteresting_descriptions'] = df_discards['Eliminate_uninteresting_descriptions'].astype(str).str.strip()
    df_discards['DELETE_(Y/N)'] = df_discards['DELETE_(Y/N)'].astype(str).str.strip()
    df_discards['NEW'] = df_discards['NEW'].astype(str).str.strip()
    df_discards_n = df_discards[df_discards['DELETE_(Y/N)'] == 'n']
    df_discards_y = df_discards[df_discards['NEW'] == 'y']

    # Note
    """
    'df_discards_n' is functionality from the previous application.
    It is currently unused but has been kept for possible future use.
    """

    return df_discards_y, df_discards_n

def forecast_process (df_forecast):
    """Normalize forecast columns and calculate the rounded-up mean quantity."""

    df_forecast = df_forecast[['PNT', 'Price_€', 'Info5', 'Info6']]
    
    df_forecast.rename(columns={'PNT':'PN'}, inplace=True)
    df_forecast['PN'] = df_forecast['PN'].astype(str).str.strip()
    df_forecast['Info5'] = pd.to_numeric(df_forecast['Info5'], errors='coerce')
    df_forecast['Info6'] = pd.to_numeric(df_forecast['Info6'], errors='coerce')
    df_forecast['Price_€'] = pd.to_numeric(df_forecast['Price_€'], errors='coerce')
    df_forecast.rename(columns={'Price_€':'Cost_Forecast(€)',}, inplace=True)


    df_forecast['Qty_Forecast'] = np.ceil((df_forecast['Info5'] + df_forecast['Info6']) / 2)    # Average rounded up across both years

    df_forecast = df_forecast[['PN', 'Qty_Forecast', 'Cost_Forecast(€)',]]

    return df_forecast


def ATA_process (df_ata_description):
    """Trim ATA codes and descriptions for consistent mapping."""

    df_ata_description['ATA'] = df_ata_description['ATA'].astype(str).str.strip()
    df_ata_description['ATA_Description'] = df_ata_description['ATA_Description'].astype(str).str.strip()

    return df_ata_description

def rule1_process (df_rule1):
    """Select and normalize the Rule 1 document and title fields."""

    df_rule1 = df_rule1[['Document', 'Title']]

    df_rule1.rename(columns={'Document':'Info_Rule1'}, inplace=True)
    df_rule1.rename(columns={'Title':'Description_Rule1'}, inplace=True)
    df_rule1['Info_Rule1'] = df_rule1['Info_Rule1'].astype(str).str.strip()
    df_rule1['Description_Rule1'] = df_rule1['Description_Rule1'].astype(str).str.strip()

    return df_rule1

def Rule2_process (df_pricing):
    """Aggregate lead time and minimum order quantity by part number."""

    df_pricing = df_pricing[['P_Material', 'P-Qtymin', 'P-PlEn']]
    df_pricing.rename(columns={'P_Material':'PN'}, inplace=True)
    df_pricing['P-Qtymin'] = df_pricing['P-Qtymin'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
    df_pricing['P-Qtymin'] = pd.to_numeric(df_pricing['P-Qtymin'], errors='coerce')
    df_pricing['P-PlEn'] = pd.to_numeric(df_pricing['P-PlEn'], errors='coerce')

    result = []
    for PN, group in df_pricing.groupby('PN'):

        lt = group['P-PlEn'].mean()
        cant = group['P-Qtymin'].mean()

        result.append({
            'PN': PN,
            'LT(days)_Rule2': lt,
            'Min_order_qty_Rule2': cant
        })

        # Same as SAP

    return pd.DataFrame(result)




# PlaneA

def Locationdata_process (df_locationdata):
    """Keep the normalized part-number-to-Rule3 mapping fields."""

    df_locationdata = df_locationdata[['PN', 'Rule3']]

    df_locationdata['PN'] = df_locationdata['PN'].astype(str).str.strip()
    df_locationdata['Rule3'] = df_locationdata['Rule3'].astype(str).str.strip()

    return df_locationdata

def merge_per_content (df, df2, col_id='PN_Raw_Material', colnorm='Info_Rule1'):

    # Build a colnorm -> Info_Rule1/Description_Rule1/etc. mapping.
    # For each Info_Rule1 value, find the PN_Raw_Material that contains it.
    map_colnorm = df2.set_index(colnorm)  # index = colnorm, columnas = col2, col3

    def buscar_colnorm(id_valor, min_len=5):
        """Return the colnorm value contained in colID, or None."""
        if pd.isna(id_valor) or str(id_valor).strip() == '':
            return None

        id_str = str(id_valor)

        for c1 in map_colnorm.index:
            if pd.isna(c1) or str(c1).strip() == '':
                continue
            c1_str = str(c1)
            if len(c1_str) >= min_len and (c1_str in id_str or id_str in c1_str):
                return c1
                
        return None

    df = df.copy()
    df['_colnorm_match'] = df[col_id].apply(buscar_colnorm)

    # Standard merge using the mapped key
    df_resuPlaneBdo = (
        df.merge(df2, left_on='_colnorm_match', right_on=colnorm, how='left')
        .drop(columns='_colnorm_match')
        .reset_index(drop=True)
    )

    return df_resuPlaneBdo

def design_data_PlaneA_process (df_design_PlaneA):
    """Normalize PlaneA design fields and convert weight from kilograms to grams."""

    df_design_PlaneA = df_design_PlaneA[['NUMBER', 'MATERIAL_PARTNUMBER', 'K1', 'WEIGHT(kg)', 'LENGTH', 'WIDTH', 'DIAM', 'THICKNESS', 'F_CLASS', 'DOMESTIC_TITLE']]
    df_design_PlaneA.rename(columns={'NUMBER':'PN'}, inplace=True)
    df_design_PlaneA.rename(columns={'MATERIAL_PARTNUMBER':'PN_Raw_Material'}, inplace=True)
    df_design_PlaneA.rename(columns={'WEIGHT(kg)':'Weight(g)'}, inplace=True)
    df_design_PlaneA.rename(columns={'LENGTH':'Length'}, inplace=True)
    df_design_PlaneA.rename(columns={'WIDTH':'Width'}, inplace=True)
    df_design_PlaneA.rename(columns={'DIAM':'Diameter'}, inplace=True)
    df_design_PlaneA.rename(columns={'THICKNESS':'Thickness'}, inplace=True)
    df_design_PlaneA.rename(columns={'F_CLASS':'F_Class'}, inplace=True)
    df_design_PlaneA.rename(columns={'DOMESTIC_TITLE':'Description_Design'}, inplace=True)

    df_design_PlaneA = df_design_PlaneA[['PN', 'PN_Raw_Material', 'K1', 'Weight(g)', 'Length', 'Width', 'Diameter', 'Thickness', 'F_Class', 'Description_Design']]

    df_design_PlaneA['Weight(g)'] = pd.to_numeric(df_design_PlaneA['Weight(g)'], errors='coerce')
    df_design_PlaneA['Weight(g)'] = df_design_PlaneA['Weight(g)'] * 1000                                # Convert to kg
    df_design_PlaneA['Length'] = pd.to_numeric(df_design_PlaneA['Length'], errors='coerce')
    df_design_PlaneA['Width'] = pd.to_numeric(df_design_PlaneA['Width'], errors='coerce')
    df_design_PlaneA['Diameter'] = pd.to_numeric(df_design_PlaneA['Diameter'], errors='coerce')
    df_design_PlaneA['Thickness'] = pd.to_numeric(df_design_PlaneA['Thickness'], errors='coerce')

    return df_design_PlaneA

def SAP_archive_process (df_SAP):
    """Aggregate SAP order history by part number."""

    df_SAP = df_SAP[['Number_material', 'Last_Mod', 'Group', 'price', 'Name_Prov', 'Qty']]

    df_SAP.rename(columns={'Number_material':'PN'}, inplace=True)
    df_SAP.rename(columns={'Last_Mod':'Date1'}, inplace=True)
    df_SAP['Date1'] = pd.to_datetime(df_SAP['Date1'], format='mixed', dayfirst=True, errors='coerce')
    df_SAP['price'] = pd.to_numeric(df_SAP['price'], errors='coerce')
    df_SAP['Qty'] = pd.to_numeric(df_SAP['Qty'], errors='coerce')

    result = []
    for PN, group in df_SAP.groupby('PN'):
        group = group.sort_values('Date1')

        order = len(group)
        qty = np.ceil(group['Qty'].mean())
        qty_sum = group['Qty'].sum()
        cost = group['price'].mean()
        aog = 'Y' if(group['Group'].astype(str).str.startswith(('X1', 'Y1'), na=False)).any() else ''
        supplier = group.iloc[-1]['Name_Prov']

        result.append({
            'PN': PN,
            'Order_Received(4Y)_SAP': order,
            'Average_Qty_per_Order_SAP': qty,
            'Sum_Qty_Units_SAP': qty_sum,
            'Average_Cost(€)_SAP': cost,
            'AOG_SAP': aog,
            'Supplier_SAP': supplier
        })

    # Aggregate the SAP history into one row per part number.

    return pd.DataFrame(result)

def Valor_stock_process (df_valor_stock):
    """Sum stock quantities and values by part number."""

    df_valor_stock = df_valor_stock[['Material', 'Stock', 'Stock_Price']]

    df_valor_stock.rename(columns={'Material':'PN'}, inplace=True)
    df_valor_stock['Stock'] = pd.to_numeric(df_valor_stock['Stock'], errors='coerce')
    df_valor_stock['Stock_Price'] = pd.to_numeric(df_valor_stock['Stock_Price'], errors='coerce')

    result = []
    for PN, group in df_valor_stock.groupby('PN'):

        stock = group['Stock'].sum()
        stock_value = group['Stock_Price'].sum()

        result.append({
            'PN': PN,
            'Stock': stock,
            'Stock_Value': stock_value
        })

    # Same as SAP

    return pd.DataFrame(result)

def Rule2_process_PlaneA (df_PlaneA_pricing):
    """Calculate average PlaneA material cost by part number."""

    df_PlaneA_pricing = df_PlaneA_pricing[['MATERIAL', 'R_Cost(EUR)', 'R_STATUS']]
    df_PlaneA_pricing = df_PlaneA_pricing[df_PlaneA_pricing['R_STATUS'].isin(['V', 'M'])]
    df_PlaneA_pricing.rename(columns={'MATERIAL':'PN'}, inplace=True)
    df_PlaneA_pricing['R_Cost(EUR)'] = pd.to_numeric(df_PlaneA_pricing['R_Cost(EUR)'], errors='coerce')
    
    result = []
    for PN, group in df_PlaneA_pricing.groupby('PN'):

        cost = group['R_Cost(EUR)'].mean()

        result.append({
            'PN': PN,
            'Cost(€)_Rule2': cost
        })

    # Same as SAP

    return pd.DataFrame(result)




# Planes B and C
# Function mapping PlaneB_code -> ATA
def relacion_ATA_PlaneB_process (df_relation_ata_PlaneB):

    df_relation_ata_PlaneB['PlaneB_code'] = df_relation_ata_PlaneB['PlaneB_code'].astype(str).str.strip()
    df_relation_ata_PlaneB['ATA'] = df_relation_ata_PlaneB['ATA'].astype(str).str.strip()

    return df_relation_ata_PlaneB


def design_PlaneB_PlaneC_process (df_design_PlaneB_PlaneC):
    """Normalize design data and expand dimension strings into separate fields."""

    if df_design_PlaneB_PlaneC is df_design_PlaneB:
        df_design_PlaneB_PlaneC = df_design_PlaneB_PlaneC[['PART_NUMBER', 'PRIME_MAT', 'MAT_SPEC', 'DESIGN_WEIGHT', 'UNIT', 'DIMENSION']]
    elif df_design_PlaneB_PlaneC is df_design_PlaneC:
        df_design_PlaneB_PlaneC = df_design_PlaneB_PlaneC[['PART_NUMBER', 'TIPO_PARTE', 'PRIME_MAT', 'MAT_SPEC', 'DESIGN_WEIGHT', 'UNIT', 'DIMENSION']]
        df_design_PlaneB_PlaneC.rename(columns={'TIPO_PARTE':'Part_Type'}, inplace=True)
        df_design_PlaneB_PlaneC['Part_Type'] = df_design_PlaneB_PlaneC['Part_Type'].astype(str).str.strip()

    df_design_PlaneB_PlaneC.rename(columns={'PART_NUMBER':'PN'}, inplace=True)
    df_design_PlaneB_PlaneC.rename(columns={'PRIME_MAT':'PN_Raw_Material'}, inplace=True)
    df_design_PlaneB_PlaneC.rename(columns={'MAT_SPEC':'RM_Norm_PlaneB/PlaneC'}, inplace=True)
    df_design_PlaneB_PlaneC.rename(columns={'DESIGN_WEIGHT':'Weight(g)'}, inplace=True)
    df_design_PlaneB_PlaneC['PN'] = df_design_PlaneB_PlaneC['PN'].astype(str).str.strip()
    df_design_PlaneB_PlaneC['PN_Raw_Material'] = df_design_PlaneB_PlaneC['PN_Raw_Material'].astype(str).str.strip()
    df_design_PlaneB_PlaneC['UNIT'] = df_design_PlaneB_PlaneC['UNIT'].astype(str).str.strip()
    df_design_PlaneB_PlaneC['DIMENSION'] = df_design_PlaneB_PlaneC['DIMENSION'].astype(str).str.strip()
    df_design_PlaneB_PlaneC['Weight(g)'] = pd.to_numeric(df_design_PlaneB_PlaneC['Weight(g)'], errors='coerce')

    mask_kg = df_design_PlaneB_PlaneC['UNIT'].str.upper() == 'KG'
    df_design_PlaneB_PlaneC.loc[mask_kg, 'Weight(g)'] = df_design_PlaneB_PlaneC.loc[mask_kg, 'Weight(g)'] * 1000

    

    # Parse the column
    def parse_norm_PlaneB_PlaneC_suffix(PlaneB_PlaneC_norm_value):
        """
        Parses the suffix in the 'RM_Norm_PlaneB/PlaneC' column to extract Diameter/Thickness.
        Two possible patterns:
        1. '...ROD<number>' at the end (hyphen before it optional) -> Diameter only
        2. '###-<number>' or '###-<number>x<number>' (suffix after last '-')
        Returns (diameter, thickness) or (None, None) if not applicable.
        """
        if pd.isna(PlaneB_PlaneC_norm_value):
            return None, None

        s = str(PlaneB_PlaneC_norm_value).strip()

        # Pattern 1: ROD<number> at the end, hyphen before it optional
        m = re.search(r'ROD\s*([0-9]\.?[0-9]+)\s$', s, flags=re.IGNORECASE)
        if m:
            return float(m.group(1)), None

        # Pattern 2: suffix after the last '-'
        if '-' not in s:
            return None, None

        suffix = s.rsplit('-', 1)[-1].strip()
        parts = re.split(r'[xX]', suffix)

        # Reject if it contains letters (e.g. 'S230') -> not useful dimensions
        if any(re.search(r'[A-Za-z]', p) for p in parts):
            return None, None

        if not all(re.fullmatch(r'[0-9]*\.?[0-9]+', p.strip()) for p in parts):
            return None, None

        nums = [float(p.strip()) for p in parts]

        if len(nums) == 1:
            # single number -> Diameter only
            return nums[0], None
        elif len(nums) == 2:
            nums.sort()
            thickness, diameter = nums  # min -> Thickness, max -> Diameter
            return diameter, thickness
        else:
            # unexpected number of parts -> discard
            return None, None

    def parse_dimensions(value, PlaneB_PlaneC_norm_value=None):
        """
        Parses a dimension string into Length, Width, Diameter, Thickness.
        If value is 'L=x', also looks at PlaneB_PlaneC_norm_value (the 'RM_Norm_PlaneB/PlaneC' column) to fill
        Diameter/Thickness from the suffix after the last '-'.
        """
        length = width = diameter = thickness = None

        if pd.isna(value) or str(value).strip() == '':
            return length, width, diameter, thickness

        s = str(value).strip()
        s = re.sub(r'\s+MM$', '', s, flags=re.IGNORECASE).strip()

        # Case: L=<value>
        m = re.fullmatch(r'[Ll]\s*=\s*([0-9]*\.?[0-9]+)', s)
        if m:
            length = float(m.group(1))
            diameter, thickness = parse_norm_PlaneB_PlaneC_suffix(PlaneB_PlaneC_norm_value)
            return length, width, diameter, thickness

        # Case: NxN or NxNxN
        parts = re.split(r'[xX]', s)
        if all(re.fullmatch(r'[0-9]*\.?[0-9]+', p.strip()) for p in parts):
            nums = sorted([float(p.strip()) for p in parts])
            if len(nums) == 2:
                width, length = nums
            elif len(nums) == 3:
                thickness, width, length = nums
            return length, width, diameter, thickness

        return length, width, diameter, thickness

    def expand_dimensions(df, dim_col='DIMENSION', norm_PlaneB_PlaneC_col='RM_Norm_PlaneB/PlaneC'):
        """
        Expands dim_col into Length, Width, Diameter, Thickness.
        Uses norm_PlaneB_PlaneC_col only for rows where dim_col is 'L=x'.
        """
        parsed = df.apply(
            lambda row: pd.Series(
                parse_dimensions(row[dim_col], row.get(norm_PlaneB_PlaneC_col)),
                index=['Length', 'Width', 'Diameter', 'Thickness']
            ),
            axis=1
        )
        pos = df.columns.get_loc(dim_col)
        df = df.drop(columns=[dim_col])
        for i, new_col in enumerate(['Length', 'Width', 'Diameter', 'Thickness']):
            df.insert(pos + i, new_col, parsed[new_col])
        
        return df

    df_design_PlaneB_PlaneC = expand_dimensions(df_design_PlaneB_PlaneC)

    return df_design_PlaneB_PlaneC


def Rule2_process_PlaneB_PlaneC (df_PlaneB_PlaneC_pricing):
    """Calculate average material cost by part number for Planes B and C."""

    df_PlaneB_PlaneC_pricing = df_PlaneB_PlaneC_pricing[['MATERIAL', 'Cost_in_EUR', 'STATUS']]
    df_PlaneB_PlaneC_pricing = df_PlaneB_PlaneC_pricing[df_PlaneB_PlaneC_pricing['STATUS'].isin(['V', 'M'])]
    df_PlaneB_PlaneC_pricing.rename(columns={'MATERIAL':'PN'}, inplace=True)
    df_PlaneB_PlaneC_pricing['Cost_in_EUR'] = pd.to_numeric(df_PlaneB_PlaneC_pricing['Cost_in_EUR'], errors='coerce')
    
    result = []
    for PN, group in df_PlaneB_PlaneC_pricing.groupby('PN'):

        cost = group['Cost_in_EUR'].mean()

        result.append({
            'PN': PN,
            'Cost(€)_Rule2': cost
        })

    # Same as SAP

    return pd.DataFrame(result)

def Obsolescences_process (df_obsolescences):
    """Exclude closed or proposed-solution records and return the remaining part numbers."""

    df_obsolescences = df_obsolescences[['PN_OBS', 'STATUS']]
    df_obsolescences.rename(columns={'PN_OBS':'PN'}, inplace=True)
    df_obsolescences['PN'] = df_obsolescences['PN'].astype(str).str.strip()
    df_obsolescences['STATUS'] = df_obsolescences['STATUS'].astype(str).str.strip()

    df_obsolescences = df_obsolescences[~df_obsolescences['STATUS'].isin(['Cerrado', 'Solución Propuesta + DC'])]

    df_obsolescences = df_obsolescences[['PN']]

    return df_obsolescences


    




###############################################################################
# Assemblies
###############################################################################

"""
Extracts assembly part numbers according to their naming convention.
Incomplete and not optimized because bulk PN_conjuntos -> PN_hijo files are unavailable for PlaneA.
"""

def pn_type1_process (df_portfolio, df_missing_parts, df_obsolescences):

    def PlaneB_type1_process (df_portfolio, df_missing_parts, df_obsolescences):
        """Build the PlaneB assembly subset and preserve source metadata."""

        if DEBUG_MODE_ON:
            print("\n DEBUG ** PlaneB Process")

        
        # Normalize filters
        df_portfolio = df_portfolio[['ID', 'Status', 'PN_PlaneB_PlaneC', 'Info7', 'Info1', 'Origin']]
        df_portfolio.rename(columns={'PN_PlaneB_PlaneC':'PN'}, inplace=True)

        df_missing_parts = df_missing_parts[df_missing_parts['Program'] == 'PlaneB']
        df_missing_parts = df_missing_parts[['PN', 'Info1', 'd_missing_parts']]

        pn_portfolio = set(df_portfolio["PN"])
        pn_missing_parts = set(df_missing_parts["PN"])
        pn_obsolescences = set(df_obsolescences['PN'])

        df_PlaneB_type1 = df_portfolio.merge(df_missing_parts, on='PN', how='left', suffixes=('_portf', '_missing'))

        df_PlaneB_type1['Info1'] = df_PlaneB_type1['Info1_portf'].fillna(df_PlaneB_type1['Info1_missing'])
        df_PlaneB_type1.drop(columns=['Info1_portf', 'Info1_missing'], inplace=True)
        
        pn_only_missing = ~df_missing_parts['PN'].isin(df_portfolio['PN'])
        df_only_missing = df_missing_parts[pn_only_missing].copy()
        df_PlaneB_type1 = pd.concat([df_PlaneB_type1, df_only_missing], ignore_index=True)

        pn_only_obsolescences = ~df_obsolescences['PN'].isin(df_PlaneB_type1['PN'])
        df_only_obsolescences = df_obsolescences[pn_only_obsolescences].copy()
        df_PlaneB_type1 = pd.concat([df_PlaneB_type1, df_only_obsolescences], ignore_index=True)



        # Determine whether the record comes from the portfolio or missing-parts data.
        def data_from (pn_val):
            in_portfolio = pn_val in pn_portfolio
            in_missing = pn_val in pn_missing_parts
            in_obsolescences = pn_val in pn_obsolescences

            sources = []
            if in_portfolio:
                sources.append("Portfolio")
            if in_missing:
                sources.append("Missing_Parts")
            if in_obsolescences:
                sources.append("Obsolescences")
            
            return " + ".join(sources) if sources else "NO DATA"
        
        df_PlaneB_type1['Data_From'] = df_PlaneB_type1['PN'].map(data_from)



        # PlaneB naming-convention filter
        regex_generic = r'^\d{2}-[^/]{5}-\d{2}[^/]{2}$'                                             # These regexes are anonymized to avoid revealing the exact PlaneB part nomenclature.
        regex_generic_suffix = r'^\d{2}-[^/]{5}-\d{2}[^/]{2}-\d{2}$'
        df_PlaneB_type1 = df_PlaneB_type1[
            (df_PlaneB_type1['PN'].str.fullmatch(regex_generic)) |
            (df_PlaneB_type1['PN'].str.fullmatch(regex_generic_suffix))
        ]

        # Filters
        df_PlaneB_type1 = df_PlaneB_type1[(df_PlaneB_type1['PN'].notna()) & (df_PlaneB_type1['PN'] != '') & (df_PlaneB_type1['PN'] != 'nan')]
        df_PlaneB_type1 = df_PlaneB_type1[(df_PlaneB_type1['Info1'].notna()) & (df_PlaneB_type1['Info1'] != '') & (df_PlaneB_type1['Info1'] != 'nan')]
        #df_PlaneB_type1 = df_PlaneB_type1[(~df_PlaneB_type1['Info7'].isin(['A', 'B']))]
        #lista_origin = ['01', '03', '05', '07', '09', '11', '13']
        #escaped_origin = [re.escape(word) for word in lista_origin]
        #regex_origin = '|'.join(escaped_origin)
        #df_PlaneB_type1 = df_PlaneB_type1[~df_PlaneB_type1['Origin'].str.contains(regex_origin, na=False)]
        

        return df_PlaneB_type1
    
    def PlaneC_type1_process (df_portfolio, df_missing_parts, df_obsolescences):
        """Build the PlaneC assembly subset and preserve source metadata."""

        if DEBUG_MODE_ON:
            print("\n DEBUG ** PlaneC Process")

        
        # Normalize filters
        df_portfolio = df_portfolio[['ID', 'Status', 'PN_PlaneB_PlaneC', 'Info7', 'Info1', 'Origin']]
        df_portfolio.rename(columns={'PN_PlaneB_PlaneC':'PN'}, inplace=True)

        df_missing_parts = df_missing_parts[df_missing_parts['Program'] == 'PlaneC']
        df_missing_parts = df_missing_parts[['PN', 'Info1', 'd_missing_parts']]

        pn_portfolio = set(df_portfolio["PN"])
        pn_missing_parts = set(df_missing_parts["PN"])
        pn_obsolescences = set(df_obsolescences['PN'])

        df_PlaneC_type1 = df_portfolio.merge(df_missing_parts, on='PN', how='left', suffixes=('_portf', '_missing'))

        df_PlaneC_type1['Info1'] = df_PlaneC_type1['Info1_portf'].fillna(df_PlaneC_type1['Info1_missing'])
        df_PlaneC_type1.drop(columns=['Info1_portf', 'Info1_missing'], inplace=True)
        
        pn_only_missing = ~df_missing_parts['PN'].isin(df_portfolio['PN'])
        df_only_missing = df_missing_parts[pn_only_missing].copy()
        df_PlaneC_type1 = pd.concat([df_PlaneC_type1, df_only_missing], ignore_index=True)

        pn_only_obsolescences = ~df_obsolescences['PN'].isin(df_PlaneC_type1['PN'])
        df_only_obsolescences = df_obsolescences[pn_only_obsolescences].copy()
        df_PlaneC_type1 = pd.concat([df_PlaneC_type1, df_only_obsolescences], ignore_index=True)



        # Determine whether the record comes from the portfolio or missing-parts data.
        def data_from (pn_val):
            in_portfolio = pn_val in pn_portfolio
            in_missing = pn_val in pn_missing_parts
            in_obsolescences = pn_val in pn_obsolescences

            sources = []
            if in_portfolio:
                sources.append("Portfolio")
            if in_missing:
                sources.append("Missing_Parts")
            if in_obsolescences:
                sources.append("Obsolescences")
            
            return " + ".join(sources) if sources else "NO DATA"
        
        df_PlaneC_type1['Data_From'] = df_PlaneC_type1['PN'].map(data_from)



        # Naming-convention filter
        regex_generic_alpha = r'[A-Z].{3}[A-Z].*'
        regex_generic_group_1 = r'([A-Z]{4}|[A-Z]{2}-).*'
        regex_generic_group_2 = r'(?![A-Z]{3})([A-Z]{2}).*'
        regex_generic_prefix = r'\d{2}-([A-Z]{2}|[A-Z]{4}).*'
        regex_generic_group_3 = r'([A-Z]{5}|[A-Z]{3}|[A-Z]{2}-|[A-Z]{3}-|[A-Z]{4}-).*'
        regex_generic_group_4 = r'([A-Z]{2}|[A-Z]{3}).*'
        df_PlaneC_type1 = df_PlaneC_type1[
            (df_PlaneC_type1['PN'].str.fullmatch(regex_generic_alpha)) |
            (df_PlaneC_type1['PN'].str.fullmatch(regex_generic_group_1)) |
            (df_PlaneC_type1['PN'].str.fullmatch(regex_generic_group_2)) |
            (df_PlaneC_type1['PN'].str.fullmatch(regex_generic_prefix)) |
            (df_PlaneC_type1['PN'].str.fullmatch(regex_generic_group_3)) |
            (df_PlaneC_type1['PN'].str.fullmatch(regex_generic_group_4))
        ]

        # Filters
        df_PlaneC_type1 = df_PlaneC_type1[(df_PlaneC_type1['PN'].notna()) & (df_PlaneC_type1['PN'] != '') & (df_PlaneC_type1['PN'] != 'nan')]
        df_PlaneC_type1 = df_PlaneC_type1[(df_PlaneC_type1['Info1'].notna()) & (df_PlaneC_type1['Info1'] != '') & (df_PlaneC_type1['Info1'] != 'nan')]
        df_PlaneC_type1 = df_PlaneC_type1[(~df_PlaneC_type1['Info7'].isin(['A', 'B']))]
        df_PlaneC_type1 = df_PlaneC_type1[(df_PlaneC_type1['Part_Type'].isin(['X', 'Y']))]
        #lista_origin = ['02', '04', '06', '08', '10', '12', '14']
        #escaped_origin = [re.escape(word) for word in lista_origin]
        #regex_origin = '|'.join(escaped_origin)
        #df_PlaneC_type1 = df_PlaneC_type1[~df_PlaneC_type1['Origin'].str.contains(regex_origin, na=False)]
        

        return df_PlaneC_type1
    
    
    df_PlaneB_type1 = PlaneB_type1_process (df_portfolio, df_missing_parts, df_obsolescences)
    df_PlaneC_type1 = PlaneC_type1_process (df_portfolio, df_missing_parts, df_obsolescences)
    
    
    return df_PlaneB_type1, df_PlaneC_type1







###############################################################################
# COMMON
###############################################################################
"""
Remove or retain certain part numbers based on discard data.
"""
def COMMON_discards (df, df_discards_y, df_discards_n):

    if DEBUG_MODE_ON:
        print("\n DEBUG ** COMMON Discards")

    # Discards
    list_y = df_discards_y['Eliminate_uninteresting_descriptions'].tolist()
    list_n = df_discards_n['Eliminate_uninteresting_descriptions'].tolist()
    escaped_y = [re.escape(word) for word in list_y]
    escaped_n = [re.escape(word) for word in list_n]
    regex_y = '|'.join(escaped_y)
    regex_n = '|'.join(escaped_n)
    df = df[(~df['Info1'].str.contains(regex_y))]


    return df

"""
- Add design data.
- Create a material dictionary to populate the RMAbr column based on matching
  description keywords (this can and should be optimized through testing).
"""
def Add_Design (df, df_design, df_rule1):

    if DEBUG_MODE_ON:
        print("\n DEBUG ** Add Design")

    # PlaneA
    if df is df_PlaneA:
        df['auxiliar'] = df['PN'].str[:-2]         # In PlaneA, remove the last two digits to compare with the design table (PASS).
        df = df.merge(df_design, left_on='auxiliar', right_on='PN', how='left', suffixes=('', '_df2'))
        df = df.drop(columns=['auxiliar', 'PN_df2'])
        df = df.drop_duplicates()
        df = merge_per_content (df, df_rule1)

    # PlaneB
    elif df is df_PlaneB or df is df_PlaneC:
        df['auxiliar'] = df['PN'].str.replace(r'-regex\d$', '', regex=True)        # Remove Rule4 for the merge while keeping the original part number with Rule4.
        df = df.merge(df_design, left_on='auxiliar', right_on='PN', how='left', suffixes=('', '_df2'))
        df = df.drop(columns=['auxiliar', 'PN_df2'])
        df = df.drop_duplicates()
        df['RM_Norm_PlaneB_aux'] = df['RM_Norm_PlaneB/PlaneC'].str.replace('-', '', regex=False).str.replace('/', '', regex=False).str.replace(' ', '', regex=False).str.strip()
        df_rule1['RM_Rule1_aux'] = df_rule1['Info_Rule1'].str.replace('-', '', regex=False).str.replace('/', '', regex=False).str.replace(' ', '', regex=False).str.strip()
        df = merge_per_content (df, df_rule1, col_id='RM_Norm_PlaneB_aux', colnorm='RM_Rule1_aux')
        df.drop(columns=['RM_Norm_PlaneB_aux', 'RM_Rule1_aux'])
        #df = df.merge(df_rule1, on='Info_Rule1', how='left')

    else:
        df['auxiliar'] = df['PN']



    df['RMAbr'] = ''

    MATERIALS = {
        'aluminium': {
            'output': 'Aluminium',
            'patterns': [r'\baluminium\b', r'\baluminum\b', r'\baluminio\b', r'\bal\b'],
            'exclusion': []
        },
        'titanium': {
            'output': 'Titanium',
            'patterns': [r'\btitanium\b', r'\btitanio\b', r'\bti\b'],
            'exclusion': []
        },
        'aramid_fibre': {
            'output': 'AF',
            'patterns': [r'\baramid\b', r'\baramid_fibre\b', r'\baf\b'],
            'exclusion': []
        },
        'cotton': {
            'output': 'Cotton',
            'patterns': [r'\bcotton\b', r'\balgodon\b' r'\bcott\b'],
            'exclusion': []
        },
        'copper': {
            'output': 'Copper',
            'patterns': [r'\bcopper\b', r'\bcobre\b', r'\bcu\b'],
            'exclusion': []
        },
        'elastomer': {
            'output': 'Elastomer',
            'patterns': [r'\belastomer\b', r'\belastomero\b', r'\bel\b'],
            'exclusion': []
        },
        'polyamide': {
            'output': 'Polyamide',
            'patterns': [r'\bpolyamide\b', r'\bPA\b', r'\bpoliamida\b'],
            'exclusion': []
        },
        'polyarylentherketone': {
            'output': 'PAEK',
            'patterns': [r'\bpolyarylentherketone\b', r'\bpoliarilentercetona \b', r'\bpaek\b'],
            'exclusion': []
        },
        'polycarbonate': {
            'output': 'Polycarbonate',
            'patterns': [r'\bpolycarbonate\b', r'\bpolicarbonato\b', r'\bpc\b'],
            'exclusion': []
        },
        'methacrylate': {
            'output': 'Methacrylate',
            'patterns': [r'\bmethacrylate\b', r'\bmetacrilato\b', r'\bpmma\b'],
            'exclusion': []
        },
        'teflon': {
            'output': 'Teflon',
            'patterns': [r'\bteflon\b', r'\bptfe\b'],
            'exclusion': []
        },
        'polyurethane': {
            'output': 'Polyurethane',
            'patterns': [r'\bpolyurethane\b', r'\bpoliuretano\b', r'\bpu\b'],
            'exclusion': []
        },
        'polyvinylchloride': {
            'output': 'PVC',
            'patterns': [r'\bpolyvinylchloridw\b', r'\bpolicloruro\b', r'\bpvc\b'],
            'exclusion': []
        },
        'steel': {
            'output': 'Steel',
            'patterns': [r'\bsteel\b', r'\bacero\b', r'\bst\b'],
            'exclusion': []
        },
        'tungsten': {
            'output': 'Tungsten',
            'patterns': [r'\btungsten\b', r'\btungsteno\b', r'\bwolfram\b', r'\bwolframio\b', r'\btu\b', r'\bw\b'],
            'exclusion': []
        },
        'other_metals': {
            'output': 'Other_Metals',
            'patterns': [r'\bmetal\b', r'\bother_metal\b', r'\bother_metals\b', r'\bmetals\b'],
            'exclusion': []
        },
        'inconel': {
            'output': 'Inconel',
            'patterns': [r'\binconel\b', r'\binc\b'],
            'exclusion': []
        },
        'polyetherimide': {
            'output': 'Ultem(PEI)',
            'patterns': [r'\bultem\b', r'\bult\b', r'\bpolyetherimide\b', r'\bpei\b', r'\bpolieterimida'],
            'exclusion': []
        },
    }

    # Use only RM data (Maria's PASS list).
    def material_classification (df, primary_col='Description_Rule1', col_search=None, col_out='RMAbr', separation=', '):
        """Classify materials from description keywords and write the results to col_out."""
        
        col_1 = df[primary_col].fillna('').astype(str).str.strip()

        if col_search:
            col_2 = df[list(col_search)].fillna('').astype(str).agg(''.join, axis=1).str.strip()
            text = col_1.where(col_1 != '', other=col_2).str.lower()
        else:
            text = col_1.str.lower()  # Use only the primary column; keep it empty if it is empty.


        results = [[] for _ in range(len(df))]

        for name, rule in MATERIALS.items():
            match_pos = text.str.contains('|'.join(rule['patterns']), flags=re.IGNORECASE, regex=True)

            if rule['exclusion']:
                match_exc = text.str.contains('|'.join(rule['exclusion']), flags=re.IGNORECASE, regex=True)
            else:
                match_exc = pd.Series(False, index=df.index)
            
            mask = match_pos & ~match_exc

            for idx in df.index[mask]:
                results[idx].append(rule['output'])
        
        df[col_out] = [separation.join(r) if r else None for r in results]

        return df

    df = material_classification (df)

    
    return df


"""
Merge a DataFrame into the main DataFrame. Used in several places.
"""
def merge_process (df, df_2):

    df = df.merge(df_2, on='PN', how='left')
    
    return df


"""
Determine whether an item is an ALM study candidate.
This section handles the 'ALM Candidate' and 'Study' columns.
The 'RULES' series identifies the logic handled automatically,
which can be easily modified, removed, or extended.
"""
def Candidate_Study (df):

    if DEBUG_MODE_ON:
        print("\n DEBUG ** Candidate Study")

    df['eval_material'] = ''
    df['eval_dimension'] = '' 
    df['eval_comercial'] = ''
    df['Candidate_X'] = ''
    df['Study'] = ''

    MATERIAL_DIMENSION = {
        'Aluminium':{'Length':50, 'Width':50, 'Thickness':50, 'Diameter':50},                           # Dimensions anonymized to avoid revealing exact PlaneA part information.
        'Steel':{'Length':50, 'Width':50, 'Thickness':50, 'Diameter':50},
        'Iconel':{'Length':50, 'Width':50, 'Thickness':50, 'Diameter':50},
        'Titanium':{'Length':50, 'Width':50, 'Thickness':50, 'Diameter':50},
        'Ultem(PEI)':{'Length':50, 'Width':50, 'Thickness':50, 'Diameter':50},
        'Polyamide':{'Length':100000, 'Width':100000, 'Thickness':100000, 'Diameter':100000}
    }
    
    def eval_material(row):
        val = row['RMAbr']
        if pd.isna(val) or str(val).strip() == '':
            return 'NO DATA'
        
        val_str = str(val).strip()

        if any(mat in val_str for mat in MATERIAL_DIMENSION):
            return 'OK'

        return 'NOK'

    def extract_material (val_str):
        for mat in MATERIAL_DIMENSION:
            if mat in val_str:
                return mat
        return None
        
    def eval_dimension(row):
        val_str = str(row['RMAbr']).strip()
        mat = extract_material (val_str)
        dims = MATERIAL_DIMENSION.get(mat, {})

        if not dims:
            return 'NO DATA'

        def have_dimension (col):
            return not (pd.isna(row[col]) or str(row[col]).strip() == '')

        have_length = have_dimension ('Length')
        have_width = have_dimension ('Width')
        have_thickness = have_dimension ('Thickness')
        have_diameter = have_dimension ('Diameter')

        classic_rute = have_length and have_width and have_thickness
        diameter_rute = have_length and have_diameter

        if classic_rute:
            cols_eval = ['Length', 'Width', 'Thickness']
        elif diameter_rute:
            cols_eval = ['Length', 'Diameter']
        else:
            presents_cols = [c for c in ['Length', 'Width', 'Thickness', 'Diameter'] if have_dimension(c)]
            if any(row[c] >= dims[c] for c in presents_cols):
                return 'NOK'
            return 'NO DATA'

        # NOK if any dimension exceeds its limit.
        if any(row[c] >= dims[c] for c in cols_eval):
            return 'NOK'

        return 'OK'


    COMERCIAL_LIMITS = {
    'Col_cost(cascade)': 300,
    'LT(days)_Rule2':    365,
    }

    def eval_comercial(row):
        val_cost_cascade = None
        for col in ['Cost(€)_Rule2', 'Average_Cost(€)_SAP', 'Cost_Forecast(€)']:
            if not pd.isna(row[col]) and str(row[col]).strip() != '':
                val_cost_cascade = row[col]
                break
        
        val_LT = row['LT(days)_Rule2']

        if val_cost_cascade is None or pd.isna(val_LT) or str(val_LT).strip() == '':
            return 'NO DATA'
        
        if val_cost_cascade >= COMERCIAL_LIMITS['Col_cost(cascade)'] or val_LT >= COMERCIAL_LIMITS['LT(days)_Rule2']:
            return 'NOK'

        return 'OK'


    ORDER_F_CLASS = [1, 2.1, 2.2, 2.3, 3.0]  # From most to least restrictive.
    MATERIAL_F_MIN = {
        'Poliamida': 3.0,
        'Al':        2.3,
        'Ti':        2.2,
        'ULTEM':     2.1,                          # F_Class anonymized to avoid revealing the exact information.
        'Inconel':   2.1,
    }
    def eval_func_class (row):
        val_str = str(row['RMAbr']).strip()
        mat = extract_material (val_str)

        # If the material has no functional restriction, this does not apply.
        if mat not in MATERIAL_F_MIN:
            return 'OK'

        val = row['F_Class']
        if pd.isna(val) or str(val).strip() == '':
            return 'NO DATA'

        funcion_min = MATERIAL_F_MIN[mat]
        funcion_row = float(val)

        # Accept if the index in the ordering is >= the required minimum index.
        idx_min = ORDER_F_CLASS.index(funcion_min)
        idx_row = ORDER_F_CLASS.index(funcion_row) if funcion_row in ORDER_F_CLASS else -1

        if idx_row == -1:
            return 'NO DATA'  # Unknown value.

        return 'OK' if idx_row >= idx_min else 'NOK'

    def evaluar_candidate(row):
        mat = eval_material(row)
        dim = eval_dimension(row)
        com = eval_comercial(row)

        # F_Class override.
        if eval_func_class(row) == 'NOK':
            return 'NO FC', 'F_CLASS not admited', mat, dim, com

        # 1. Dimensions NOK -> NO (highest priority).
        if dim == 'NOK':
            return 'NO', 'Dimensions over the limits', mat, dim, com

        # 2. Material and dimensions NO DATA.
        if mat == 'NO DATA':
            return 'NO DATA (MAT)', 'Missing MATERIAL data', mat, dim, com

        # 3. Dimensions NO DATA with a known material.
        if dim == 'NO DATA' and mat != 'NO DATA':
            #if eval_func_class(row) == 'NO DATA':
                #return 'NO DATA (DIM, FC)', 'Not enough DIMENSION data. Missing F_CLASS data', mat, dim, com
            return 'NO DATA (DIM)', 'Not enough DIMENSION data', mat, dim, com

        # 4. Material NO DATA with dimensions OK.

        # 5. Material NOK with dimensions OK -> for analysis.
        if mat == 'NOK':
            return 'TO BE ANALYSED (MAT)', 'MATERIAL known, but not in the Roadmap', mat, dim, com

        # 6. Material OK and dimensions OK -> depends on commercial evaluation.
        if mat == 'OK' and dim == 'OK':
            if com == 'OK':
                #if eval_func_class(row) == 'NO DATA':
                    #return 'YES (NO DATA FC)', 'Missing F_CLASS data', mat, dim, com
                return 'YES', '', mat, dim, com
            elif com == 'NOK':
                #if eval_func_class(row) == 'NO DATA':
                    #return 'YES (NOK €, NO DATA FC)', 'MATERIAL and DIMENSION ok, COMMERCIAL data under limits. Missing F_CLASS data', mat, dim, com
                return 'YES (NOK €)', 'MATERIAL and DIMENSION ok, COMMERCIAL data under limits', mat, dim, com
            else:  # NO DATA
                #if eval_func_class(row) == 'NO DATA':
                    #return 'YES (NO DATA €, NO DATA FC)', 'MATERIAL and DIMENSION ok, not enough COMMERCIAL data. Missing F_CLASS data', mat, dim, com
                return 'YES (NO DATA €)', 'MATERIAL and DIMENSION ok, not enough COMMERCIAL data', mat, dim, com

        # Fallback
        return 'NOT EVALUATED', '', mat, dim, com


    
    df[['Candidate_X', 'Study', 'eval_material', 'eval_dimension', 'eval_comercial']] = df.apply(lambda row: pd.Series(evaluar_candidate(row)), axis=1)

    return df




###############################################################################
# PlaneA
###############################################################################

"""
- Combine portfolio and missing-parts data and record each item's source.
- Filter usable part numbers using the provided naming convention and regex.
"""
def PlaneA_process (df_portfolio, df_missing_parts):

    if DEBUG_MODE_ON:
        print("\n DEBUG ** PlaneA Process")

    df_portfolio = df_portfolio[['ID', 'Status', 'PN_PlaneA', 'Info4', 'Info1']]
    df_portfolio.rename(columns={'PN_PlaneA':'PN'}, inplace=True)
    df_portfolio = df_portfolio[df_portfolio['Status'] != 'DRAFT']
    df_portfolio = df_portfolio[df_portfolio['Info4'] != 0]

    df_missing_parts = df_missing_parts[df_missing_parts['Program'] == 'PlaneA']
    df_missing_parts = df_missing_parts[['PN', 'Info1', 'd_missing_parts']]

    pn_portfolio = set(df_portfolio["PN"])
    pn_missing_parts = set(df_missing_parts["PN"])

    df_PlaneA = df_portfolio.merge(df_missing_parts, on='PN', how='left', suffixes=('_portf', '_missing'))

    df_PlaneA['Info1'] = df_PlaneA['Info1_portf'].fillna(df_PlaneA['Info1_missing'])
    df_PlaneA.drop(columns=['Info1_portf', 'Info1_missing'], inplace=True)
    
    pn_only_missing = ~df_missing_parts['PN'].isin(df_portfolio['PN'])
    df_only_missing = df_missing_parts[pn_only_missing].copy()
    df_PlaneA = pd.concat([df_PlaneA, df_only_missing], ignore_index=True)

    
    # Determine whether the record comes from the portfolio or missing-parts data.
    def data_from (pn_val):
        in_portfolio = pn_val in pn_portfolio
        in_missing = pn_val in pn_missing_parts
        if in_portfolio and in_missing:
            return "Portfolio + Missing_Parts"
        elif in_portfolio:
            return "Portfolio"
        else:
            return "Missing_parts"
    
    df_PlaneA['Data_From'] = df_PlaneA['PN'].map(data_from)



    # Naming-convention filter
    regex_generic_14 = r'^[A-Z][^-/][^-/][^-/][^-/]{6}[^-/]{4}$'
    regex_generic_14_suffix = r'^[A-Z][^-/][^-/][^-/][^-/]{6}[^-/]{4}-\d{2}$'
    regex_generic_prefix_3 = r'^[A-Z]{3}'
    regex_generic_prefix_5 = r'^[A-Z]{5}'

    df_PlaneA = df_PlaneA[
        (df_PlaneA['PN'].str.fullmatch(regex_generic_14)) |
        (df_PlaneA['PN'].str.fullmatch(regex_generic_14_suffix)) |
        (df_PlaneA['PN'].str.match(regex_generic_prefix_3)) |
        (df_PlaneA['PN'].str.match(regex_generic_prefix_5))
    ]

    # Filters
    df_PlaneA = df_PlaneA[df_PlaneA['PN'].notna() & (df_PlaneA['PN'] != '') & (df_PlaneA['PN'] != 'nan')]


    return df_PlaneA




"""Assign ATA using the two digits after Regex.
ATA is not assigned in special cases.
"""
def PlaneA_ATA (df_PlaneA, df_ata_description, df_locationdata):

    if DEBUG_MODE_ON:
        print("\n DEBUG ** PlaneA ATA")

    def get_code(id_val):
        if pd.isna(id_val):
            return None
        if re.match(r'^Regex[A-Za-z]', id_val):
            prefix = id_val[:5]
            match = df_locationdata[df_locationdata['PN'].str.contains(prefix, na=False)]
            if not match.empty:
                return str(match.iloc[0]['Rule3'])[:2]
            return None
        elif id_val.startswith('TERM1'):
            if len(id_val) >= 13:
                return id_val[11:13]
            return None
        else:
            m = re.match(r'^Regex(\d{2})', id_val)
            return m.group(1) if m else None

    df_PlaneA['ATA'] = df_PlaneA['PN'].apply(get_code)
    mapa = df_ata_description.set_index('ATA')['ATA_Description']
    df_PlaneA['ATA_Description'] = df_PlaneA['ATA'].map(mapa)

    

    if DEBUG_MODE_ON:
        print(f"Number of rows in df_PlaneA: {df_PlaneA.shape[0]}")
        print(f"Number of columns in df_PlaneA: {df_PlaneA.shape[1]}")
    

    return df_PlaneA







"""
The PlaneB, PlaneC, and BOOM section is not yet complete.
"""
###############################################################################
# PlaneB
###############################################################################

def PlaneB_process (df_portfolio, df_missing_parts, df_obsolescences):
    """Combine PlaneB sources, annotate origins, and filter part numbers."""

    if DEBUG_MODE_ON:
        print("\n DEBUG ** PlaneB Process")

    
    # Normalize filters
    df_portfolio = df_portfolio[['ID', 'Status', 'PN_PlaneB_PlaneC', 'Info7', 'Info1', 'Origin']]
    df_portfolio.rename(columns={'PN_PlaneB_PlaneC':'PN'}, inplace=True)

    df_missing_parts = df_missing_parts[df_missing_parts['Program'] == 'PlaneB']
    df_missing_parts = df_missing_parts[['PN', 'Info1', 'd_missing_parts']]

    pn_portfolio = set(df_portfolio["PN"])
    pn_missing_parts = set(df_missing_parts["PN"])
    pn_obsolescences = set(df_obsolescences['PN'])

    df_PlaneB = df_portfolio.merge(df_missing_parts, on='PN', how='left', suffixes=('_portf', '_missing'))

    df_PlaneB['Info1'] = df_PlaneB['Info1_portf'].fillna(df_PlaneB['Info1_missing'])
    df_PlaneB.drop(columns=['Info1_portf', 'Info1_missing'], inplace=True)
    
    pn_only_missing = ~df_missing_parts['PN'].isin(df_portfolio['PN'])
    df_only_missing = df_missing_parts[pn_only_missing].copy()
    df_PlaneB = pd.concat([df_PlaneB, df_only_missing], ignore_index=True)

    pn_only_obsolescences = ~df_obsolescences['PN'].isin(df_PlaneB['PN'])
    df_only_obsolescences = df_obsolescences[pn_only_obsolescences].copy()
    df_PlaneB = pd.concat([df_PlaneB, df_only_obsolescences], ignore_index=True)



    # Determine whether the record comes from the portfolio or missing-parts data.
    def data_from (pn_val):
        in_portfolio = pn_val in pn_portfolio
        in_missing = pn_val in pn_missing_parts
        in_obsolescences = pn_val in pn_obsolescences

        sources = []
        if in_portfolio:
            sources.append("Portfolio")
        if in_missing:
            sources.append("Missing_Parts")
        if in_obsolescences:
            sources.append("Obsolescences")
        
        return " + ".join(sources) if sources else "NO DATA"
    
    df_PlaneB['Data_From'] = df_PlaneB['PN'].map(data_from)



    # Naming-convention filter
    regex_generic = r'^\d{2}-[^/]{5}-[^/]{4}$'
    regex_generic_suffix = r'^\d{2}-[^/]{5}-[^/]{4}-\d{2}$'
    regex_generic_alpha_prefix = r'([A-Z0-9]{5}|[A-Z]{3}|[A-Z]{2}-|[A-Z]{3}-|[A-Z]{4}-).*'
    regex_generic_alpha_group = r'([A-Z]{2}|[A-Z]{3}).*'
    regex_generic_numeric_group = r'\d{2}-\d{2}.*'
    df_PlaneB = df_PlaneB[
        (df_PlaneB['PN'].str.fullmatch(regex_generic)) |
        (df_PlaneB['PN'].str.fullmatch(regex_generic_suffix)) |
        (df_PlaneB['PN'].str.fullmatch(regex_generic_alpha_prefix)) |
        (df_PlaneB['PN'].str.fullmatch(regex_generic_alpha_group))
    ]

    # Filters
    df_PlaneB = df_PlaneB[(df_PlaneB['PN'].notna()) & (df_PlaneB['PN'] != '') & (df_PlaneB['PN'] != 'nan')]
    df_PlaneB = df_PlaneB[(df_PlaneB['Info1'].notna()) & (df_PlaneB['Info1'] != '') & (df_PlaneB['Info1'] != 'nan')]
    df_PlaneB = df_PlaneB[(~df_PlaneB['Info7'].isin(['A', 'B']))]
    #lista_origin = ['03', '04', '06', '07', '09', '10', '11', '12']
    #escaped_origin = [re.escape(word) for word in lista_origin]
    #regex_origin = '|'.join(escaped_origin)
    #df_PlaneB = df_PlaneB[~df_PlaneB['Origin'].str.contains(regex_origin, na=False)]
    

    return df_PlaneB








###############################################################################
# PlaneC & BOOM
###############################################################################

def PlaneC_process (df_portfolio, df_missing_parts, df_obsolescences):
    """Combine PlaneC sources, annotate origins, and filter part numbers."""

    if DEBUG_MODE_ON:
        print("\n DEBUG ** PlaneC Process")

    
    # Normalize filters
    df_portfolio = df_portfolio[['ID', 'Status', 'PN_PlaneB_PlaneC', 'Info7', 'Info1', 'Origin']]
    df_portfolio.rename(columns={'PN_PlaneB_PlaneC':'PN'}, inplace=True)

    df_missing_parts = df_missing_parts[df_missing_parts['Program'] == 'PlaneC']
    df_missing_parts = df_missing_parts[['PN', 'Info1', 'd_missing_parts']]

    pn_portfolio = set(df_portfolio["PN"])
    pn_missing_parts = set(df_missing_parts["PN"])
    pn_obsolescences = set(df_obsolescences['PN'])

    df_PlaneC = df_portfolio.merge(df_missing_parts, on='PN', how='left', suffixes=('_portf', '_missing'))

    df_PlaneC['Info1'] = df_PlaneC['Info1_portf'].fillna(df_PlaneC['Info1_missing'])
    df_PlaneC.drop(columns=['Info1_portf', 'Info1_missing'], inplace=True)
    
    pn_only_missing = ~df_missing_parts['PN'].isin(df_portfolio['PN'])
    df_only_missing = df_missing_parts[pn_only_missing].copy()
    df_PlaneC = pd.concat([df_PlaneC, df_only_missing], ignore_index=True)

    pn_only_obsolescences = ~df_obsolescences['PN'].isin(df_PlaneC['PN'])
    df_only_obsolescences = df_obsolescences[pn_only_obsolescences].copy()
    df_PlaneC = pd.concat([df_PlaneC, df_only_obsolescences], ignore_index=True)



    # Determine whether the record comes from the portfolio or missing-parts data.
    def data_from (pn_val):
        in_portfolio = pn_val in pn_portfolio
        in_missing = pn_val in pn_missing_parts
        in_obsolescences = pn_val in pn_obsolescences

        sources = []
        if in_portfolio:
            sources.append("Portfolio")
        if in_missing:
            sources.append("Missing_Parts")
        if in_obsolescences:
            sources.append("Obsolescences")
        
        return " + ".join(sources) if sources else "NO DATA"
    
    df_PlaneC['Data_From'] = df_PlaneC['PN'].map(data_from)



    # Naming-convention filter
    regex_generic_alpha = r'[A-Z].{3}[A-Z].*'
    regex_generic_group_1 = r'([A-Z]{4}|[A-Z]{2}-).*'
    regex_generic_group_2 = r'(?![A-Z]{3})([A-Z]{2}).*'
    regex_generic_prefix = r'\d{2}-([A-Z]{2}|[A-Z]{4}).*'
    regex_generic_group_3 = r'([A-Z0-9]{5}|[A-Z]{3}|[A-Z]{2}-|[A-Z]{3}-|[A-Z]{4}-).*'
    regex_generic_group_4 = r'([A-Z]{2}|[A-Z]{3}).*'
    df_PlaneC = df_PlaneC[
        (df_PlaneC['PN'].str.fullmatch(regex_generic_alpha)) |
        (df_PlaneC['PN'].str.fullmatch(regex_generic_group_1)) |
        (df_PlaneC['PN'].str.fullmatch(regex_generic_group_2)) |
        (df_PlaneC['PN'].str.fullmatch(regex_generic_prefix)) |
        (df_PlaneC['PN'].str.fullmatch(regex_generic_group_3)) |
        (df_PlaneC['PN'].str.fullmatch(regex_generic_group_4))
    ]

    # Filters
    df_PlaneC = df_PlaneC[(df_PlaneC['PN'].notna()) & (df_PlaneC['PN'] != '') & (df_PlaneC['PN'] != 'nan')]
    df_PlaneC = df_PlaneC[(df_PlaneC['Info1'].notna()) & (df_PlaneC['Info1'] != '') & (df_PlaneC['Info1'] != 'nan')]
    df_PlaneC = df_PlaneC[(~df_PlaneC['Info7'].isin(['A', 'B']))]
    df_PlaneC = df_PlaneC[(df_PlaneC['Part_Type'].isin(['X', 'Y']))]
    #lista_origin = ['01', '02', '03', '04', '05', '06', '07', '08', '09']
    #escaped_origin = [re.escape(word) for word in lista_origin]
    #regex_origin = '|'.join(escaped_origin)
    #df_PlaneC = df_PlaneC[~df_PlaneC['Origin'].str.contains(regex_origin, na=False)]
    

    return df_PlaneC

def ATA_PlaneB (df_PlaneB, df_ata_description, df_relation_ata_PlaneB):
    """Extract PlaneB codes from part numbers and map them to ATA descriptions."""

    if DEBUG_MODE_ON:
        print("\n DEBUG ** PlaneB ATA")

    ATA_extraido = df_PlaneB['PN'].str.extract(r'^(?:regex_pattern)-(\d{2})')
    df_PlaneB['PlaneB_code'] = ATA_extraido[0]
    mapa_PlaneB = df_relation_ata_PlaneB.set_index('PlaneB_code')['ATA']
    df_PlaneB['ATA'] = df_PlaneB['PlaneB_code'].map(mapa_PlaneB)
    mapa = df_ata_description.set_index('ATA')['ATA_Description']
    df_PlaneB['ATA_Description'] = df_PlaneB['ATA'].map(mapa)
    

    if DEBUG_MODE_ON:
        print(f"Number of rows in df_PlaneB: {df_PlaneB.shape[0]}")
        print(f"Number of columns in df_PlaneB: {df_PlaneB.shape[1]}")

    
    return df_PlaneB

def PlaneC_ATA (df_PlaneC, df_ata_description, df_locationdata):
    """Extract or look up ATA codes and map them to descriptions."""

    if DEBUG_MODE_ON:
        print("\n DEBUG ** PlaneC ATA")

    def get_code(id_val):
        if pd.isna(id_val):
            return None
        if re.match(r'^regexx[A-Za-z]', id_val):
            prefix = id_val[:5]
            match = df_locationdata[df_locationdata['PN'].str.contains(prefix, na=False)]
            if not match.empty:
                return str(match.iloc[0]['Rule3'])[:2]
            return None
        elif id_val.startswith('TERM1'):
            if len(id_val) >= 14:
                return id_val[12:14]
            return None
        elif re.match(r'^regex(\d{2})', id_val):
            m = re.match(r'^regex(\d{2})', id_val)
            return m.group(1) if m else None
        elif re.match(r'^regexpattern(\d{2})', id_val):
            m = re.match(r'^regexpattern(\d{2})', id_val)            
            return m.group(1) if m else None
        else:
            m = re.match(r'^F(\d{2})', id_val)
            return m.group(1) if m else None

    df_PlaneC['ATA'] = df_PlaneC['PN'].apply(get_code)
    mapa = df_ata_description.set_index('ATA')['ATA_Description']
    df_PlaneC['ATA_Description'] = df_PlaneC['ATA'].map(mapa)

    

    if DEBUG_MODE_ON:
        print(f"Number of rows in df_PlaneC: {df_PlaneC.shape[0]}")
        print(f"Number of columns in df_PlaneC: {df_PlaneC.shape[1]}")
    

    return df_PlaneC





###############################################################################
#%% Main Body
###############################################################################
"""
This is the main section, where the functions defined above are called.
The order of execution matters because some operations depend on others.
For example, df_portfolio must be normalized before it can be used.

First, the user is asked whether to run the function that returns a file
containing all assembly parts. The user then provides the corresponding input file.
As noted above, this section is incomplete and may need a full redesign because
there is no bulk PlaneA assembly-to-child-part-number extract. For now, it returns
a file containing all assemblies that meet the naming convention in Portfolio
and Missing_Parts.

If the user chooses to run the assembly section, they are then asked whether to
run the full Screening process. If so, the complete process runs, starting with
file normalization and ending with the automated study and Excel output.
"""

ejecuted_type1 = False # Flag for when the assembly section is enabled.
proceso = input(
    """


    
    Type the number of the function to execute:
    1 - P_Type_1 (Get the P_Type_1 PN)
    2 - Screening (Main Part)

""").strip()
while True:
    if proceso == '1':
        df_portfolio = portfolio_process (df_portfolio)
        df_missing_parts = missing_parts_process (df_missing_parts)
        df_obsolescences = Obsolescences_process (df_obsolescences)
        df_PlaneB_type1, df_PlaneC_type1 = pn_type1_process (df_portfolio, df_missing_parts, df_obsolescences)

        # Create the Excel file with the final DataFrame.
        output_file = os.path.join(main_path, 'output', 'P_Type_1.xlsx')
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_PlaneB_type1.to_excel(writer, sheet_name='PlaneB_type1', index=False)
            df_PlaneC_type1.to_excel(writer, sheet_name='PlaneC_type1', index=False)
        
        ejecutar = input(
        """
        
        Remember to add the file "P_Type_1_List.xlsx" to the 'inputs' folder (deleting the old one) before continue

        Do you want to continue to the execution of the Screening (main part)? (Y/N)
        

        """).strip().upper()

        if ejecutar == 'Y':
            ejecuted_type1 = True
            proceso = '2'
            continue
        elif ejecutar == 'N':
            t001 = time.time()       
            print("""
        
        Process completed in """ + str(int((t001 - t000))) + """ s
        """)

            break
        else:
            print('No valid option, Screening will not be executed')

    elif proceso == '2':

        # NORMALIZE
        # Common
        if not ejecuted_type1:
            df_portfolio = portfolio_process (df_portfolio)
            df_missing_parts = missing_parts_process (df_missing_parts)
            df_obsolescences = Obsolescences_process (df_obsolescences)
        df_discards_y, df_discards_n = descartes_process (df_discards)
        df_rule1 = rule1_process (df_rule1)
        df_ata_description = ATA_process (df_ata_description)
        df_forecast = forecast_process (df_forecast)
        df_pricing = Rule2_process (df_pricing)
        
        #PlaneA
        df_design_PlaneA = design_data_PlaneA_process (df_design_PlaneA)
        df_locationdata = Locationdata_process (df_locationdata)
        df_PlaneA_pricing = Rule2_process_PlaneA (df_PlaneA_pricing)
        df_SAP = SAP_normalize_process (df_SAP, skip_rows=7)
        df_SAP = SAP_archive_process (df_SAP)
        db_creator(db_SAP_path, "df_SAP", df_SAP)
        
        # Planes B and C
        df_relation_ata_PlaneB = relacion_ATA_PlaneB_process (df_relation_ata_PlaneB)
        df_design_PlaneB = design_PlaneB_PlaneC_process (df_design_PlaneB)
        df_design_PlaneC = design_PlaneB_PlaneC_process (df_design_PlaneC)
        df_PlaneB_PlaneC_pricing = Rule2_process_PlaneB_PlaneC (df_PlaneB_PlaneC_pricing)


        # PROCESS
        # PlaneA
        df_PlaneA = PlaneA_process (df_portfolio, df_missing_parts)
        df_PlaneA = COMMON_discards (df_PlaneA, df_discards_y, df_discards_n)
        df_PlaneA = PlaneA_ATA (df_PlaneA, df_ata_description, df_locationdata)
        df_PlaneA = Add_Design (df_PlaneA, df_design_PlaneA, df_rule1)
        df_PlaneA = merge_process (df_PlaneA, df_PlaneA_pricing)
        df_PlaneA = merge_process (df_PlaneA, df_pricing)
        df_PlaneA = merge_process (df_PlaneA, df_SAP)
        df_PlaneA = merge_process (df_PlaneA, df_forecast)
        df_PlaneA = Candidate_Study (df_PlaneA)
        df_PlaneA.rename(columns={'K1':'Part_Type'}, inplace=True)           # Rename column K1 here to avoid issues.

        df_PlaneA = df_PlaneA[['PN', 'Info4', 'Info1', 'Part_Type', 'Data_From', 'ATA', 'ATA_Description', 'PN_Raw_Material', 'Info_Rule1', 'RMAbr', 'Weight(g)', 'Length', 'Width', 'Diameter', 'Thickness', 'F_Class', 'Cost(€)_Rule2', 'LT(days)_Rule2', 'Min_order_qty_Rule2', 'Order_Received(4Y)_SAP', 'Average_Qty_per_Order_SAP', 'Sum_Qty_Units_SAP', 'Average_Cost(€)_SAP', 'AOG_SAP', 'Qty_Forecast', 'Cost_Forecast(€)', 'd_missing_parts', 'Candidate_X', 'Study']]

        """duplicates = df_PlaneA[df_PlaneA.duplicated(subset=['PN'], keep=False)]
        print(f"Total duplicate rows: {len(duplicates)}")
        duplicates_det = duplicates[duplicates['Part_Type'] == 'DevDetPart']
        print(f"Duplicate detail-part rows: {len(duplicates_det)}")"""


        # PlaneB
        df_PlaneB = PlaneB_process (df_portfolio, df_missing_parts, df_obsolescences)
        df_PlaneB = COMMON_discards (df_PlaneB, df_discards_y, df_discards_n)
        df_PlaneB = ATA_PlaneB (df_PlaneB, df_ata_description, df_relation_ata_PlaneB)
        df_PlaneB = Add_Design (df_PlaneB, df_design_PlaneB, df_rule1)
        df_PlaneB = merge_process (df_PlaneB, df_PlaneB_PlaneC_pricing)
        df_PlaneB = merge_process (df_PlaneB, df_pricing)
        df_PlaneB = merge_process (df_PlaneB, df_SAP)
        df_PlaneB = merge_process (df_PlaneB, df_forecast)
        df_PlaneB = Candidate_Study (df_PlaneB)
        
        df_PlaneB = df_PlaneB[['PN', 'Info7', 'Info1', 'Origin', 'Data_From', 'ATA', 'ATA_Description', 'PN_Raw_Material', 'RM_Norm_PlaneB/PlaneC', 'Info_Rule1', 'Description_Rule1', 'RMAbr', 'Weight(g)', 'Length', 'Width', 'Diameter', 'Thickness', 'Cost(€)_Rule2', 'LT(days)_Rule2', 'Min_order_qty_Rule2', 'Order_Received(4Y)_SAP', 'Average_Qty_per_Order_SAP', 'Sum_Qty_Units_SAP', 'Average_Cost(€)_SAP', 'AOG_SAP', 'Qty_Forecast', 'Cost_Forecast(€)', 'd_missing_parts', 'Candidate_X', 'Study']]

        #PlaneC
        df_PlaneC = PlaneC_process (df_portfolio, df_missing_parts, df_obsolescences)
        df_PlaneC = COMMON_discards (df_PlaneC, df_discards_y, df_discards_n)
        df_PlaneC = PlaneC_ATA (df_PlaneC, df_ata_description, df_locationdata)
        df_PlaneC = Add_Design (df_PlaneC, df_design_PlaneC, df_rule1)
        df_PlaneC = merge_process (df_PlaneC, df_PlaneB_PlaneC_pricing)
        df_PlaneC = merge_process (df_PlaneC, df_pricing)
        df_PlaneC = merge_process (df_PlaneC, df_SAP)
        df_PlaneC = merge_process (df_PlaneC, df_forecast)
        df_PlaneC = Candidate_Study (df_PlaneC)

        df_PlaneC = df_PlaneC[['PN', 'Info7', 'Info1', 'Part_Type', 'Origin', 'Data_From', 'ATA', 'ATA_Description', 'PN_Raw_Material', 'RM_Norm_PlaneB/PlaneC', 'Info_Rule1', 'Description_Rule1', 'RMAbr', 'Weight(g)', 'Length', 'Width', 'Diameter', 'Thickness', 'Cost(€)_Rule2', 'LT(days)_Rule2', 'Min_order_qty_Rule2', 'Order_Received(4Y)_SAP', 'Average_Qty_per_Order_SAP', 'Sum_Qty_Units_SAP', 'Average_Cost(€)_SAP', 'AOG_SAP', 'Qty_Forecast', 'Cost_Forecast(€)', 'd_missing_parts', 'Candidate_X', 'Study']]

        # Create the Excel file with the final DataFrame.
        output_file = os.path.join(main_path, 'output', 'Screening.xlsx')
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_PlaneA.to_excel(writer, sheet_name='PlaneA', index=False)
            df_PlaneB.to_excel(writer, sheet_name='PlaneB', index=False)
            df_PlaneC.to_excel(writer, sheet_name='PlaneC', index=False)
        

        t001 = time.time()       
        print("""
        
        Process completed in """ + str(int((t001 - t000))) + """ s
        """)

        break
    else:
        print('No valid option')
        break
    