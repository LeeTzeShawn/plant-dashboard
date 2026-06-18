# %%
import pandas as pd
import numpy as np
import re
# from IPython.display import display
from collections import defaultdict

# %%
# Show all rows
pd.set_option('display.max_rows', None)

# Show all columns
pd.set_option('display.max_columns', None)

# Prevent column width truncation
pd.set_option('display.max_colwidth', None)

# Optional: Adjust display width for wide DataFrames
pd.set_option('display.width', None)

pd.options.display.float_format = '{:.6f}'.format

# %%
file_path = r"C:\WFH\MOSB_Costing_Mar26.xlsm"

# plantM1 = pd.read_excel(file_path, sheet_name="M1_Plant_Overh", engine='openpyxl')
# plantM3 = pd.read_excel(file_path, sheet_name="M3_Plant_Overh", engine='openpyxl') 
# plantM4 = pd.read_excel(file_path, sheet_name="M4_Plant_Overh", engine='openpyxl')
# plantM5 = pd.read_excel(file_path, sheet_name="M5_Plant_Overh", engine='openpyxl')
# plantM6 = pd.read_excel(file_path, sheet_name="M6_Plant_Overh", engine='openpyxl')
# plantW1 = pd.read_excel(file_path, sheet_name="W1_Plant_Overh2", engine='openpyxl') #Failed
# plantCBS = pd.read_excel(file_path, sheet_name="CBS_Plant_Overh", engine='openpyxl')
# plantCBS2 = pd.read_excel(file_path, sheet_name="CBS2_Plant_Overh", engine='openpyxl')
# plantHydro = pd.read_excel(file_path, sheet_name="Hydro_Plant_Overh", engine='openpyxl')
# plantHydro2 = pd.read_excel(file_path, sheet_name="Hydro2_Plant_Overh", engine='openpyxl')
# plantFract = pd.read_excel(file_path, sheet_name="Fract_Plant_Overh", engine='openpyxl') 
# plantIE = pd.read_excel(file_path, sheet_name="IE_Plant_Overh", engine='openpyxl') #RESUME PlantIE before work

# %%
def alternating_nan_with_prev_rule(df):
    # Step 1: convert 0 → NaN
    df = df.replace(0, np.nan)

    # Step 2: detect empty columns
    empty_mask = df.isna().all(axis=0).values
    cols = df.columns

    result_cols = []
    keep_flag = False  # controls alternating behavior

    for i in range(len(cols)):
        # ✅ NEW RULE: always keep 2nd column if it is empty
        if i == 1 and empty_mask[i]:
            result_cols.append((cols[i], df.iloc[:, i]))
            keep_flag = False  # reset pattern
            continue

        if empty_mask[i]:
            # If previous column is NOT empty → reset pattern
            if i == 0 or not empty_mask[i - 1]:
                keep_flag = False  # first NaN → drop

            if keep_flag:
                result_cols.append((cols[i], df.iloc[:, i]))

            # flip pattern
            keep_flag = not keep_flag

        else:
            # reset when hitting non-NaN column
            keep_flag = False
            result_cols.append((cols[i], df.iloc[:, i]))

    return pd.DataFrame({k: v for k, v in result_cols})

# %%
def clean_columns(df):
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()                 # remove leading/trailing spaces
        .str.replace(r'\s*-\s*$', '', regex=True)  # remove trailing " -"
        .str.strip()
    )
    return df

def clean_key(col):
    return col.astype(str).str.strip().str.upper()

# %%
def get_ALL_UNCLEANED_product_names(plant_name, plant_key=None):
    get_row_consumption = lambda x:x.astype(str).str.contains("Consumption")
    result = plant_name.apply(get_row_consumption).any(axis=1)  

    rows = result[result].index
    header_row = rows[1]
    header_row_df = plant_name.iloc[[header_row]]
    header_row_df = header_row_df.iloc[1:]
    plant_name = plant_name.loc[header_row:128, :].reset_index(drop=True)
    plant_name.columns = plant_name.iloc[0]
    
    plant_name.columns = plant_name.columns.astype(str).str.strip().str.replace('\xa0', ' ', regex=False)
    plant_name = plant_name.loc[:, plant_name.columns.notna()]
    plant_name = plant_name.loc[1:, :'TOTAL']
    plant_name = plant_name.iloc[:, 1:-1]
    plant_name = clean_columns(plant_name)
    all_uncleaned_cols = list(plant_name.columns)

    remove_words_list = ['PURCHASES']
    all_uncleaned_cols = [item for item in all_uncleaned_cols if item not in remove_words_list]

    return all_uncleaned_cols

# %%
def get_product_names(plant_name, plant_key=None):
    get_row_consumption = lambda x:x.astype(str).str.contains("Consumption")
    result = plant_name.apply(get_row_consumption).any(axis=1)  

    rows = result[result].index
    header_row = rows[1]
    header_row_df = plant_name.iloc[[header_row]]
    header_row_df = header_row_df.iloc[1:]
    plant_name = plant_name.loc[header_row:128, :].reset_index(drop=True)
    plant_name.columns = plant_name.iloc[0]
    
    plant_name.columns = plant_name.columns.astype(str).str.strip().str.replace('\xa0', ' ', regex=False)
    plant_name = plant_name.loc[:, plant_name.columns.notna()]
    plant_name = plant_name.loc[1:, :'TOTAL']
    plant_name = plant_name.iloc[:, 1:-1]
    plant_name = clean_columns(plant_name)
    
    drop_col_names = ['PURCHASES', '0', 'ALL', 'TOTAL(PMF)', 'TOTAL(H)'] #'0' also included in the list
    for col in drop_col_names:
        if plant_name.columns.isin([col]).any():
            plant_name = plant_name.drop(columns=col)
        else:
            pass
    
    cols = list(plant_name.columns)
    plant_name.columns = cols
    product_names = plant_name.columns[1:] 
    product_names = product_names.str.strip()
    product_names = product_names[product_names.notna()]
    product_names = list(product_names)
  
    def find_duplicates_with_positions(lst):
        positions = defaultdict(list)
        for i, val in enumerate(lst):
            positions[val].append(i)
        return {k: v for k, v in positions.items() if len(v) > 1}

    # 🔍 Step 1: Detect duplicates
    duplicates = find_duplicates_with_positions(product_names)
        
    return product_names

# %%
RENAME_MAP = {
    "plantM3": {
        3: "RBDHST_REWORK_(M3)",
        6: "RBDHPKO_Tolling",   # index of duplicate to rename
        37: "BHPO_1MAX_(REWORK)",
        38: "HSBO",
    },

    "plantM4": {
        19: 'RBDPKOL-M4_2'
    },
    
    "plantM5": {
        2: "RBDPO_REWORK_2"
    },
    
    "plantM6": {
        17: "RBDHPO_4648",
        18: "RBDHPO_4648_REWORK",
        21: "DRBDCBR",
        36: "RBDIEFAT_012_Tolling",
    },

    "plantHydro" : {
        9: "CHPKOL_(REWORK)",
        14: "CHPKOL_1MAX_3",
    },

    "plantHydro2" : {
        0: "CHPKST",
        4: "CHPKOL_(HYDRO2)_2",
        8: "CHPKOL_(HYDRO2)_3",
    },

    }

def rename_duplicates(lst, plant_key=None):
    if plant_key in RENAME_MAP:
        for idx, new_name in RENAME_MAP[plant_key].items():
            if idx < len(lst):
                lst[idx] = new_name
    return lst

# %%
def get_ConsumptionList(plant_name, plant_key=None):
    get_row_consumption = lambda x:x.astype(str).str.contains("Consumption")
    result_row_consumption = plant_name.apply(get_row_consumption, axis=1).any(axis=1)

    get_row_description = lambda x:x.astype(str).str.contains("Description")
    result_row_description = plant_name.apply(get_row_description, axis=1).any(axis=1)

    header_row = result_row_consumption[result_row_consumption].index[1]
    final_row = result_row_description[result_row_description].index[0]
    
    consumption_df = plant_name.iloc[header_row:final_row-1, :].reset_index(drop=True)
    consumption_df.columns = consumption_df.loc[0]
    consumption_df = clean_columns(consumption_df)
    
    consumption_df = consumption_df.loc[1:].reset_index(drop=True)
    consumption_df = consumption_df.loc[:, 'Consumption':'TOTAL']
    consumption_df = consumption_df.rename(columns={'TOTAL':'Quantity_MT_Total'})

    cols_list = consumption_df.columns.dropna().drop_duplicates().tolist()
    consumption_df = consumption_df[cols_list]
    drop_col_names = ['PURCHASES', '0', 'ALL', 'TOTAL(PMF)', 'TOTAL(H)'] #'0' also included in the list
    for col in drop_col_names:
        if consumption_df.columns.isin([col]).any():
            consumption_df = consumption_df.drop(columns=col)
        else:
            pass
    
    product_names = rename_duplicates(get_product_names(plant_name, plant_key), plant_key)

    cols = list(consumption_df.columns)
    
    consumption_df = consumption_df.rename(columns={'Quantity_MT_Total':'Quantity_Consumed_MT_Total'})
    cols = list(consumption_df.columns)
    cols[1:-1] = product_names
    consumption_df.columns = cols
    
    # Step 2: apply renaming ONLY on matching columns

    CONSUMPTION_LIST = pd.Categorical(consumption_df['Consumption'])  
    CONSUMPTION_LIST = list(dict.fromkeys(CONSUMPTION_LIST))
    
    consumption_df = pd.melt(consumption_df,
                             id_vars=['Consumption', 'Quantity_Consumed_MT_Total'],
                             value_vars=product_names,
                             var_name='Product_ID',
                             value_name='Quantity_Consumed_MT').reset_index(drop=True)
    
    consumption_df[['Quantity_Consumed_MT', 'Quantity_Consumed_MT_Total']] = consumption_df[['Quantity_Consumed_MT', 'Quantity_Consumed_MT_Total']].apply(lambda x:pd.to_numeric(x, errors='coerce').fillna(0)).reset_index(drop=True)

    return consumption_df, CONSUMPTION_LIST


# %%
def GL_Cleaner(plant_name, plant_key=None):
    consumption_df, CONSUMPTION_LIST = get_ConsumptionList(plant_name, plant_key)
    plant_name_copy = plant_name.copy()
    get_row_description = lambda x:x.astype(str).str.contains("Description") #Add back .astype(str)
    result = plant_name.apply(get_row_description, axis=1).any(axis=1)
    header_row = result[result].index[0]
    header_row_df = plant_name.iloc[[header_row]]
    header_row_df = header_row_df.iloc[1:]
    GL = plant_name.loc[header_row:127, :].reset_index(drop=True)
    GL.columns = GL.loc[0]
    GL = GL.loc[1:]
    GL = clean_columns(GL) ## Apply the clean_columns function here:
    delete_row_labels = ['ADD : DIRECT EXPENSES', 'DIRECT RAW MATERIALS', 'DIRECT LABOUR', 'DIRECT FACTORY OVERHEADS', 'CONTRACT WORKERS-PACKER', 'CONTRACT WORKERS-GENERAL', 'GRAND TOTAL', np.nan]
    GL = GL[~GL['Description'].isin(delete_row_labels)].reset_index(drop=True)
    GL = GL.drop(columns=['Cost/Mt']).reset_index(drop=True)
    GL = GL.loc[:, :'Total Cost of Plant'].iloc[:, :-1]
    GL = GL.loc[:, GL.columns.notna()]
    
    product_names = rename_duplicates(get_product_names(plant_name, plant_key), plant_key)
    all_uncleaned_product_names = get_ALL_UNCLEANED_product_names(plant_name, plant_key)
    
    cols = list(GL.columns)

    if plant_key == 'plantHydro':
        cols[4:] = product_names
        #GL.columns = cols (Check if the location of this matters)
        check_colnames = ['PMT', 'Total Cost of Plant', 'ALL', 'TOTAL(PMF)', 'TOTAL(H)'] #Added 'ALL', 'TOTAL(PMF)', 'TOTAL(H)' in the list
        GL = GL.drop(columns=check_colnames, errors='ignore').reset_index(drop=True) #Don't drop unwanted columns too early!
        GL.columns = cols

    elif plant_key == 'plantM4':
        positions = [4, 10, 11, 14, 15, 16, 17]
        for pos in sorted(positions, reverse=True):
            GL.insert(pos, 'Total Cost of Produce', 0, allow_duplicates=True)
        check_colnames = ['PMT', 'Total Cost of Plant', 'ALL', 'TOTAL(PMF)', 'TOTAL(H)'] #Added 'ALL', 'TOTAL(PMF)', 'TOTAL(H)' in the list
        GL = GL.drop(columns=check_colnames, errors='ignore').reset_index(drop=True) #Don't drop unwanted columns too early!
        cols = list(GL.columns)
        cols[4:] = product_names
        GL.columns = cols
        
    else:
        cols[4:] = all_uncleaned_product_names[1:] #product_names_renamed already removed all the ALL and TOTAL(PMF) and TOTAL(H). So this causes column mismatch.
        GL.columns = cols
        check_colnames = ['PMT', 'Total Cost of Plant', 'ALL', 'TOTAL(PMF)', 'TOTAL(H)'] #Added 'ALL', 'TOTAL(PMF)', 'TOTAL(H)' in the list
        GL = GL.drop(columns=check_colnames, errors='ignore').reset_index(drop=True) #Don't drop unwanted columns too early!
    
    product_names = rename_duplicates(get_product_names(plant_name, plant_key), plant_key)
    
    cols = list(GL.columns)
    cols[4:] = product_names
    GL.columns = cols

    if plant_key == 'plantHydro2':
        GL = GL.rename(columns={'Amount':'AMOUNT'})
    GL = pd.melt(
                GL,
                id_vars=['Description', 'Category', 'A/C CODE', 'AMOUNT'],
                value_vars=product_names,
                var_name='Product_ID',
                value_name='Amount_RM',
            )
    
    def GLCode_Map_CostCategory(gl_account):
        if str(gl_account).startswith('700500'):
            return 'Direct cost'
        elif str(gl_account).startswith('705'):
            return 'Factory overhead'
        else:
            return 'Labour cost'
    GL['Cost_Category'] = GL['A/C CODE'].apply(GLCode_Map_CostCategory)
    GL = GL.rename(columns={'AMOUNT':'Amount_RM_Total', 
                            'A/C Code':'GL_Account'
                            })
    GL[['Amount_RM_Total', 'Amount_RM']] = GL[['Amount_RM_Total', 'Amount_RM']].apply(lambda x:pd.to_numeric(x, errors='coerce').fillna(0))
    GL = GL.reset_index(drop=True)
    
    Cleaned_insideConsumptionList_ordered = []
    for word in CONSUMPTION_LIST:
        if word == 'NICKEL':
            new_word = 'Nickel Catalyst'
            Cleaned_insideConsumptionList_ordered.append(new_word)
        elif word == 'STEAM':
            new_word = 'Total Steam Cost'
            Cleaned_insideConsumptionList_ordered.append(new_word)
        elif word == 'SODIUM METHOXIDE CONSUMPTION':
            new_word = 'Sodium Methoxide'
            Cleaned_insideConsumptionList_ordered.append(new_word)
        elif word == 'ELECTRICITY':
            new_word = 'Electricity'
            Cleaned_insideConsumptionList_ordered.append(new_word)
        elif word == 'Filter Ad':
            new_word = 'Operating Supplies Expenses - Celetom Filter'
            Cleaned_insideConsumptionList_ordered.append(new_word)
        elif word == 'OP-SUPPLIES EXPENSES - OTHERS':
            new_word = 'Operating Supplies Expenses - Others'
            Cleaned_insideConsumptionList_ordered.append(new_word)
        else:
            word = word.strip().title().replace('-',' ')
            Cleaned_insideConsumptionList_ordered.append(word)

    GLItems_Inside_ConsumptionList = GL[GL['Category'].isin(Cleaned_insideConsumptionList_ordered)].reset_index(drop=True)
    ##### GLItems_Inside_ConsumptionList is NOT following the ConsumptionItem order in CONSUMPTION_LIST. Adjust the order so that it matches.#####

    #Adding items that are present in ConsumptionList, but named differently inside GL.
    if plant_key == 'plantFract':
        GLItems_Inside_ConsumptionList = GLItems_Inside_ConsumptionList.replace({'Total Steam Cost':'STEAM',
                                                                                'Operating Supplies Expenses - Celetom Filter':'Filter Ad',
                                                                                'Electricity':'ELECTRICITY'})
    elif plant_key == 'plantIE':
        # print("This is plant IE!")
        GLItems_Inside_ConsumptionList = GLItems_Inside_ConsumptionList.replace({'Sodium Methoxide':'SODIUM METHOXIDE CONSUMPTION',
                                                                                'Nitrogen Gas':'NITROGEN GAS',
                                                                                'Citric Acid':'CITRIC ACID',
                                                                                'Phosphoric Acid':'PHOSPHORIC ACID',
                                                                                'Total Steam Cost':'STEAM',
                                                                                'Electricity':'ELECTRICITY',
                                                                                'Bleaching Earth':'BLEACHING EARTH',
                                                                                'Diesel':'DIESEL',
                                                                                'Hydrogen Gas':'HYDROGEN GAS',
                                                                                'Activated Carbon':'ACTIVATED CARBON',
                                                                                'Natural Gas':'NATURAL GAS',
                                                                                'Operating Supplies Expenses - Celetom Filter':'Filter Ad',
                                                                                'Operating Supplies Expenses - Others':'OP-SUPPLIES EXPENSES - OTHERS',
                                                                                'Water': 'WATER',
                                                                                })
    else:
        GLItems_Inside_ConsumptionList = GLItems_Inside_ConsumptionList.replace({'Bleaching Earth':'BLEACHING EARTH',
                                                                                'Phosphoric Acid':'PHOSPHORIC ACID',
                                                                                'Citric Acid':'CITRIC ACID',
                                                                                'Activated Carbon':'ACTIVATED CARBON',
                                                                                'Diesel':'DIESEL',
                                                                                'Nitrogen Gas':'NITROGEN GAS',
                                                                                'Nickel Catalyst':'NICKEL',
                                                                                'Hydrogen Gas':'HYDROGEN GAS',
                                                                                'Electricity':'ELECTRICITY',
                                                                                'Natural Gas':'NATURAL GAS',
                                                                                'Total Steam Cost':'STEAM',
                                                                                'Operating Supplies Expenses - Others':'OP-SUPPLIES EXPENSES-OTHERS',
                                                                                'Operating Supplies Expenses - Celetom Filter':'Filter Ad',
                                                                                'Water': 'WATER'})
      
    GLItems_Inside_ConsumptionList['Category'] = pd.Categorical(GLItems_Inside_ConsumptionList['Category'],
                                                               categories = CONSUMPTION_LIST,
                                                                 ordered = True)
    GLItems_Inside_ConsumptionList = GLItems_Inside_ConsumptionList.sort_values('Category').reset_index(drop=True)

    exclude_words = ['COGS - B/EARTH CONSUMED', 'WATER', 'COGS - PHOSPHORIC A. CONSUMED','COGS - CITRIC ACID CONSUMED',
                     'COGS - ACT. CARBON CONSUMED','COGS - DIESEL CONSUMED','COGS - NITROGEN',
                     'COGS - NICKEL CONSUMED', 'COGS - HYDROGEN GAS', 'Electricity', 'NATURAL GAS ', 
                     'BOILER COST ALLOCATION','OPERATING SUPPLIES', 'COGS - CELATOM FILTER']
    pattern = '|'.join(exclude_words)
    
    GLItems_NotInside_ConsumptionList = GL[~GL['Description'].str.contains(pattern, case=False, na=False)].reset_index(drop=True)
    # print("Below is the GLItems_Inside_ConsumptionList: ")
    # display(GLItems_Inside_ConsumptionList)
    return GLItems_Inside_ConsumptionList, GLItems_NotInside_ConsumptionList, GL


# %%
def get_ConsumptionOutput(plant_name, plant_key=None): #Set up varoius if-scenarios for plantM3, plantM1, plantM5, plantM6...all the plants
    row_mask_CONSUMPTION = plant_name.astype(str).apply(
        lambda row: row.str.contains("Consumption", case=False, regex=False).any(),
        axis=1
    )
    first_row_index_CONSUMPTION = row_mask_CONSUMPTION.idxmax()      # row label
    first_row_pos_CONSUMPTION   = row_mask_CONSUMPTION.values.argmax()  # row position
    col_mask = plant_name.astype(str).apply(
    lambda col: col.str.contains("Consumption", case=False, regex=False).any(),
    axis=0
    )
    last_col_index_CONSUMPTION = col_mask[col_mask].index[-1]     # column name
    last_col_pos_CONSUMPTION   = col_mask.values.nonzero()[0][-1] # column position

    # print(f"The first row index of CONSUMPTION is {first_row_index_CONSUMPTION}")
    # print(f"The last column index of CONSUMPTION is {last_col_index_CONSUMPTION}")

    row_mask = plant_name.astype(str).apply(
        lambda row: row.str.contains("Total (MT)", case=True, regex=False).any(),
        axis=1
    )

    first_row_index_TotalMT = row_mask.idxmax()       # row label
    first_row_pos_TotalMT   = row_mask.values.argmax()  # row position
    # print(f"The first row index of Total (MT) is {first_row_index_TotalMT}")

    col_mask_TotalMT = plant_name.astype(str).apply(
    lambda col: col.str.contains("Total (MT)", case=True, regex=False).any(),
    axis=0
)
    # ✅ FIRST occurrence (instead of last)
    first_col_index_TotalMT = col_mask_TotalMT.idxmax()       # column name
    first_col_pos_TotalMT   = col_mask_TotalMT.values.argmax() # column position
    # print(f"The first column index of Total (MT) is: {first_col_index_TotalMT}")
    row_mask_TOTAL = plant_name.astype(str).apply(
        lambda row: row.str.contains("TOTAL", case=True, regex=False).any(),
        axis=1
    )
    first_row_index_TOTAL = row_mask_TOTAL.idxmax()        # row label
    first_row_pos_TOTAL   = row_mask_TOTAL.values.argmax() # row position
    # print(f"The first row index of TOTAL is: {first_row_index_TOTAL}")

    row_series = plant_name.iloc[first_row_pos_TOTAL]
    col_mask_TOTAL = row_series.str.contains("Total", case=False, regex=False)
    first_col_index_TOTAL = col_mask_TOTAL.idxmax()        # column name
    first_col_pos_TOTAL   = col_mask_TOTAL.values.argmax() # column position
    # print(f"The first column index of TOTAL is: {first_col_index_TOTAL}")
    result = plant_name.loc[0:first_row_index_TOTAL, first_col_index_TOTAL:first_col_index_TotalMT].reset_index(drop=True)
    
    col0 = result.iloc[:, 0].astype(str)
    pattern_ConsumptionMT = r'^consumption(?:\s*\[MT\])?$'
    start_idx = col0[col0.str.contains(pattern_ConsumptionMT, case=False, regex=True, na=False)].index[0] #Error is here
    end_idx_mask = col0.str.contains("Output", case=False, regex=False) | \
                                    col0.str.contains("Produced", case=False, regex=False)
    end_idx = col0[end_idx_mask].index[0]
         
    result = result.loc[start_idx:end_idx, :]

    cols = list(result.columns)
    cols[0] = 'Description'
    result.columns = cols
    result = result.drop(columns=result.columns[1]) 
    result = alternating_nan_with_prev_rule(result)

    uncleaned_product_names = get_ALL_UNCLEANED_product_names(plant_name, plant_key)
    # display(result)
    # print(f"This is the uncleaned product names: {uncleaned_product_names}, length is {len(uncleaned_product_names)}")
    # print(f"This is the columns of result: {result.columns}, length: {len(result.columns)}")
    cols = list(result.columns)
    if plant_key == 'plantHydro' or plant_key == 'plantHydro2':
        # display(result)
        product_names = rename_duplicates(get_product_names(plant_name, plant_key), plant_key)
        # print("This is plantHydro!")
        # print(f"This is all the renamed cleaned product names: {product_names}, length: {len(product_names)}")
        # print(f"This is all the columns in cols: {cols}, length: {len(cols)}")
        cols[1:-1] = product_names
        # print('---')
        # print(f"This is renamed duplicates product_names: {product_names}, length: {len(product_names)}")
        # print(f"This is all the GL columns: {result.columns}, length: {len(result.columns)}")
        result.columns = cols
        
    elif plant_key == 'plantW1':
        cols[2:-1] = uncleaned_product_names[1:]
        result.columns = cols
        result = result.drop(result.columns[[1, 2]], axis=1)

    elif plant_key == 'plantM4':
        # print("This is plantM4!")
        result = result.iloc[:, 2:]
        # print("Below is the result dataframe of plantM4: ")
        # display(result)
        # print(f"This is the length of the result df: {len(result.columns)}")
        # print(f"This is the uncleaned_product_names: {uncleaned_product_names}, length: {len(uncleaned_product_names)}")
        cols[:-1] = uncleaned_product_names
        result.columns = cols

    else:
        cols[:-1] = uncleaned_product_names
        result.columns = cols
    result = result.rename(columns={result.columns[-1]: 'Quantity_MT_Total',
                                    'Consumption':'Description'})
    # display(result)
    return result


def getConsumptionOnly(plant_name, plant_key = None):
    ConsumptionOutput = get_ConsumptionOutput(plant_name, plant_key)
    ConsumptionOnly = pd.DataFrame(ConsumptionOutput.iloc[0]).T
    ConsumptionOnly = pd.melt(
        ConsumptionOnly,
        id_vars = ['Description', 'Quantity_MT_Total'],
        value_vars = ConsumptionOnly[1:-1],
        var_name = 'Product_ID',
        value_name='Quantity_Consumed_MT'
    )
    ConsumptionOnly = ConsumptionOnly.rename(columns={'Quantity_MT_Total':'Quantity_Consumed_MT_Total'})
    # display(ConsumptionOnly)
    return ConsumptionOnly


def getOutputOnly(plant_name, plant_key = None):
    ConsumptionOutput = get_ConsumptionOutput(plant_name, plant_key)
    OutputOnly = pd.DataFrame(ConsumptionOutput.iloc[1]).T
    OutputOnly = pd.melt(
        OutputOnly,
        id_vars = ['Description', 'Quantity_MT_Total'],
        value_vars = OutputOnly[1:-1],
        var_name = 'Product_ID',
        value_name='Quantity_Outputted_MT'
    )
    OutputOnly = OutputOnly.rename(columns={'Quantity_MT_Total':'Quantity_Outputted_MT_Total'})
    # display(OutputOnly)
    return OutputOnly

# %%
def getDC_OH_Inside_ConsumptionList(plant_name, plant_key=None): 
    GLItems_Inside_ConsumptionList, GLItems_NotInside_ConsumptionList, GL = GL_Cleaner(plant_name, plant_key)
    
    GLItems_Inside_ConsumptionList_DC_OH = GLItems_Inside_ConsumptionList[GLItems_Inside_ConsumptionList['Cost_Category'].isin(['Direct cost', 'Factory overhead'])].reset_index(drop=True)
    
    consumption_df, CONSUMPTION_LIST = get_ConsumptionList(plant_name, plant_key)
    consumption_df['Consumption'] = clean_key(consumption_df['Consumption'])
    consumption_df['Product_ID'] = clean_key(consumption_df['Product_ID'])

    GLItems_Inside_ConsumptionList['Category'] = clean_key(GLItems_Inside_ConsumptionList['Category'])
    GLItems_Inside_ConsumptionList['Product_ID'] = clean_key(GLItems_Inside_ConsumptionList['Product_ID'])
    
    GLItems_Inside_ConsumptionList['Category'] = clean_key(GLItems_Inside_ConsumptionList['Category'])
    GLItems_Inside_ConsumptionList['Product_ID'] = clean_key(GLItems_Inside_ConsumptionList['Product_ID'])
    print("Columns in GLItems_Inside_ConsumptionList:")
    print(GLItems_Inside_ConsumptionList.columns)

    
    # ✅ Clean keys
    GLItems_Inside_ConsumptionList_DC_OH['Category'] = clean_key(GLItems_Inside_ConsumptionList_DC_OH['Category'])
    GLItems_Inside_ConsumptionList_DC_OH['Product_ID'] = clean_key(GLItems_Inside_ConsumptionList_DC_OH['Product_ID'])

    # ✅ Aggregate safely
    GL_agg = (
        GLItems_Inside_ConsumptionList_DC_OH
        .groupby(['Category', 'Product_ID', 'A/C CODE', 'Cost_Category'], as_index=False)
        ['Amount_RM_Total']
        .sum()
    )

    # ✅ Merge
    merged_df = consumption_df.merge(
        GL_agg,
        left_on=['Consumption', 'Product_ID'],
        right_on=['Category', 'Product_ID'],
        how='left'
    )

    merged_df = merged_df.drop(columns=['Category'])


    # merged_df = consumption_df.merge(
    #     GLItems_Inside_ConsumptionList[
    #         ['Category', 'Product_ID', 'Amount_RM_Total']
    #     ],
    #     left_on=['Consumption', 'Product_ID'],
    #     right_on=['Category', 'Product_ID'],
    #     how='left'
    # )
    # merged_df = merged_df.drop(columns=['Category'])
    merged_df = merged_df.sort_index()
    merged_df['Unit_Rate'] = merged_df['Amount_RM_Total'] / merged_df['Quantity_Consumed_MT_Total'].where(merged_df['Quantity_Consumed_MT_Total']>0)
    merged_df['Allocated_Cost'] = merged_df['Unit_Rate'] * merged_df['Quantity_Consumed_MT']
    # display(merged_df)
    return merged_df

def getDC_OH_NotInside_ConsumptionList(plant_name, plant_key=None):
    getOutputOnly_df = getOutputOnly(plant_name, plant_key)
    
    GLItems_Inside_ConsumptionList, GLItems_NotInside_ConsumptionList, GL = GL_Cleaner(plant_name, plant_key)
    GLItems_NotInside_ConsumptionList_DC_OH = GLItems_NotInside_ConsumptionList[GLItems_NotInside_ConsumptionList['Cost_Category'].isin(['Direct cost', 'Factory overhead'])].reset_index(drop=True)
    GLItems_NotInside_ConsumptionList_DC_OH[['Amount_RM', 'Amount_RM_Total']] = GLItems_NotInside_ConsumptionList_DC_OH[['Amount_RM', 'Amount_RM_Total']].apply(lambda x:pd.to_numeric(x, errors='coerce'))
    
    ProductID_QuantityOutputtedMTTotal_Mapping = dict(zip(getOutputOnly_df['Product_ID'], getOutputOnly_df['Quantity_Outputted_MT_Total']))
    ProductID_QuantityOutputtedMT_Mapping = dict(zip(getOutputOnly_df['Product_ID'], getOutputOnly_df['Quantity_Outputted_MT']))
    
    mapped = GLItems_NotInside_ConsumptionList_DC_OH['Product_ID'].map(
        ProductID_QuantityOutputtedMTTotal_Mapping
    )

    mapped_again = GLItems_NotInside_ConsumptionList_DC_OH['Product_ID'].map(
        ProductID_QuantityOutputtedMT_Mapping
    )

    GLItems_NotInside_ConsumptionList_DC_OH['Quantity_Outputted_MT'] = mapped_again
    GLItems_NotInside_ConsumptionList_DC_OH.loc[:, 'Quantity_Outputted_MT_Total'] = mapped
    GLItems_NotInside_ConsumptionList_DC_OH['Unit_Rate'] = GLItems_NotInside_ConsumptionList_DC_OH['Amount_RM_Total'] / GLItems_NotInside_ConsumptionList_DC_OH['Quantity_Outputted_MT_Total'].where(GLItems_NotInside_ConsumptionList_DC_OH['Quantity_Outputted_MT_Total']>0)
    GLItems_NotInside_ConsumptionList_DC_OH['Allocated_Cost'] = GLItems_NotInside_ConsumptionList_DC_OH['Unit_Rate'] * GLItems_NotInside_ConsumptionList_DC_OH['Quantity_Outputted_MT']
    GLItems_NotInside_ConsumptionList_DC_OH = GLItems_NotInside_ConsumptionList_DC_OH.drop(columns=['Amount_RM'])
    # display(GLItems_NotInside_ConsumptionList_DC_OH)
    return GLItems_NotInside_ConsumptionList_DC_OH

def getLC_NotInside_ConsumptionList(plant_name, plant_key):
    GLItems_Inside_ConsumptionList, GLItems_NotInside_ConsumptionList, GL = GL_Cleaner(plant_name, plant_key)
    GLItems_NotInside_ConsumptionList_LC = GLItems_NotInside_ConsumptionList[GLItems_NotInside_ConsumptionList['Cost_Category']=='Labour cost'].reset_index(drop=True)
    
    getOutputOnly_df = getOutputOnly(plant_name, plant_key)
    ProductID_QuantityOutputtedMTTotal_Mapping = dict(zip(getOutputOnly_df['Product_ID'], getOutputOnly_df['Quantity_Outputted_MT_Total']))
    ProductID_QuantityOutputtedMT_Mapping = dict(zip(getOutputOnly_df['Product_ID'], getOutputOnly_df['Quantity_Outputted_MT']))

    GLItems_NotInside_ConsumptionList_LC[['Amount_RM', 'Amount_RM_Total']] = GLItems_NotInside_ConsumptionList_LC[['Amount_RM', 'Amount_RM_Total']].apply(lambda x:pd.to_numeric(x, errors='coerce'))

    mapped = GLItems_NotInside_ConsumptionList_LC['Product_ID'].map(
        ProductID_QuantityOutputtedMTTotal_Mapping
    )

    mapped_again = GLItems_NotInside_ConsumptionList_LC['Product_ID'].map(
        ProductID_QuantityOutputtedMT_Mapping
    )
    
    GLItems_NotInside_ConsumptionList_LC['Quantity_Outputted_MT'] = mapped_again
    GLItems_NotInside_ConsumptionList_LC.loc[:, 'Quantity_Outputted_MT_Total'] = mapped

    GLItems_NotInside_ConsumptionList_LC['Allocation_%'] = GLItems_NotInside_ConsumptionList_LC['Quantity_Outputted_MT'] / GLItems_NotInside_ConsumptionList_LC['Quantity_Outputted_MT_Total'].where(GLItems_NotInside_ConsumptionList_LC['Quantity_Outputted_MT_Total']>0)
    GLItems_NotInside_ConsumptionList_LC['Total_Labour_Cost_Pool'] = GLItems_NotInside_ConsumptionList_LC['Amount_RM'].sum()
    GLItems_NotInside_ConsumptionList_LC['Allocated_Cost'] = GLItems_NotInside_ConsumptionList_LC['Allocation_%'] * GLItems_NotInside_ConsumptionList_LC['Total_Labour_Cost_Pool']
    
    GLItems_NotInside_ConsumptionList_LC = GLItems_NotInside_ConsumptionList_LC.reset_index(drop=True)
    # display(GLItems_NotInside_ConsumptionList_LC)
    return GLItems_NotInside_ConsumptionList_LC




# %%
def process_plants_to_excel(uploaded_file, output_file="cleaned_output.xlsx"):
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
    #Create a mapping from sheet name to the plant_key your functions expect
    plant_key_mapping = {
        'W1_Plant_Overh2': 'plantW1',
        'Fract_Plant_Overh': 'plantFract',
        'M1_Plant_Overh': 'plantM1',
        'M3_Plant_Overh': 'plantM3',
        'M4_Plant_Overh': 'plantM4',
        'M5_Plant_Overh': 'plantM5',
        'M6_Plant_Overh': 'plantM6',
        'CBS_Plant_Overh': 'plantCBS',
        'CBS2_Plant_Overh': 'plantCBS2',
        'Hydro_Plant_Overh': 'plantHydro',
        'Hydro2_Plant_Overh': 'plantHydro2',
        'IE_Plant_Overh': 'plantIE'
    }
    # failed_plants = []

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, plant_df in all_sheets.items():
            if sheet_name in plant_key_mapping:
                plant_key = plant_key_mapping[sheet_name] #Use mapped key!
                print(f"Processing {sheet_name} as '{plant_key}'...")
                    #Process ONE plant at a time
                df1 = getDC_OH_Inside_ConsumptionList(plant_df.copy(), plant_key)
                df2 = getDC_OH_NotInside_ConsumptionList(plant_df.copy(), plant_key)
                df3 = getLC_NotInside_ConsumptionList(plant_df.copy(), plant_key)

                #Save immediately
                sheet_name1 = f"{plant_key}_Inside"[:31]
                sheet_name2 = f"{plant_key}_NotInside"[:31]
                sheet_name3 = f"{plant_key}_LC"[:31]

                df1.to_excel(writer, sheet_name=sheet_name1, index=False)
                df2.to_excel(writer, sheet_name=sheet_name2, index=False)
                df3.to_excel(writer, sheet_name=sheet_name3, index=False)

                print(f" ✅ Saved {plant_key}")
                
                # except Exception as e:
                #     print(f"Failed {plant_key}: {e}")
                #     failed_plants.append(plant_key)
                #     continue
        
    #Add this summary block
    # print(f"\n{'='*50}")
    # print(f"✅ Successful: {len(valid_plants) - len(failed_plants)} plants")
    # print(f"❌Failed: {len(failed_plants)} plants - {failed_plants}")
    # print(f"\n{'='*50}")
    print(f"\nDone! Output saved to {output_file}")

### Paste new getTotalAllocatedCost function here:
# New addition: Total_Allocated_Cost = Direct Material Cost + Factory Overhead Cost + Labour Cost, but per Product_ID
def get_Total_AllocatedCost_perProduct(df_inside, df_notinside, df_lc, plant_name):
    """
    Calculate Total Allocated Cost per Product ID
    Total = Direct Material + Factory Overhead + Labour Cost
    """
    inside_costs = df_inside.groupby('Product_ID')['Allocated_Cost'].sum().reset_index()
    inside_costs.columns = ['Product_ID', 'Inside_Cost']

    notinside_costs = df_notinside.groupby('Product_ID')['Allocated_Cost'].sum().reset_index()
    notinside_costs.columns = ['Product_ID', 'NotInside_Cost']                                                      
    
    lc_costs = df_lc.groupby('Product_ID')['Allocated_Cost'].sum().reset_index()
    lc_costs.columns = ['Product_ID', 'LC_Cost']

    result = pd.merge(inside_costs, notinside_costs, on='Product_ID', how='outer')
    result = pd.merge(result, lc_costs, on='Product_ID', how='outer')
    result = result.fillna(0)

    result['Total_Allocated_Cost'] = result['Inside_Cost'] + result['NotInside_Cost'] + result['LC_Cost']

    return result.sort_values('Total_Allocated_Cost', ascending=False)



# if __name__ == '__main__':
#     file_path = r"C:\WFH\MOSB_Costing_Mar26.xlsm"
#     process_plants_to_excel(file_path)




    

