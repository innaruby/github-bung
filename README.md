import tkinter as tk
from tkinter import filedialog
import xlwings as xw
import pandas as pd
import numpy as np
import io
import re 
# ---------------------------------------------------------------------
# Creating The Product Tree
# ---------------------------------------------------------------------
def get_product_MSR_pos_alignment(pt_kst_xl:pd.ExcelFile, sheet_name:str, header_product:int=1):
    """
    Reads the specified sheet using `header_row` as the header,
    processes it, and prints the 'MSR groups -> product numbers'
    output into the output_area widget.
    """
    # Parse the desired sheet
    try: 
        prodtree = pt_kst_xl.parse(sheet_name=sheet_name,header=header_product-1)
        if 'Produkt' in prodtree.columns:
            # Remove lines where 'Produkt' is empty (NaN)
            prodtree = prodtree.dropna(subset=['Produkt'])
            # create a dict that maps "product code" -> "MSR position"
            pt = {x[1].iloc[0]:x[1].iloc[1] if pd.notnull(x[1].iloc[1]) else 'N/A' for x in prodtree.iterrows()} #pt -> product tree
        else:
            print('There is no "Produkt" column found with this header row!')
            pt = None
        return pt
    except Exception as e:
        print('Error while reading in and processing the MSR Cost centers', e)
        return None

# ---------------------------------------------------------------------
# Creating The MSR Output Format options
# ---------------------------------------------------------------------
def get_msr_output_format(pt_kst_xl:pd.ExcelFile,sheet_name:str,header_cost_center:int=1):
    try:
        ccdf = pt_kst_xl.parse(sheet_name=sheet_name,header=header_cost_center-1)
        ccs = ccdf.groupby(ccdf[['MSR','Untersegment']].apply(lambda x: ' - '.join(x.tolist()), axis=1)).agg({'KST':'unique'}).squeeze().to_dict()
    except Exception as e:
        print('Error while reading in and processing the Produktbaum', e)
        return None
    
    return ccs

# ---------------------------------------------------------------------
# Opening an excel via a dialogbox
# ---------------------------------------------------------------------
def open_an_excel(titlestr:str|None=None):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file_bytes = filedialog.askopenfile(mode='rb',title=titlestr, filetypes=[('Excel files','.xlsx')],parent=root)
    if file_bytes:
        file = pd.ExcelFile(io.BytesIO(file_bytes.read()))
        root.destroy()
        return file
    else:
        root.destroy()
        return None


# ---------------------------------------------------------------------
# Main function to creating the report
# ---------------------------------------------------------------------
def create_report(
        report_xl:pd.ExcelFile,
        gesamt_sheet_name:str,
        fg_sheet_name:str,
        current_pt:dict,
        used_cost_centers:list[int],
        header_row:int=9,
        cost_center_colname:str='Kostenstelle des Geschäfts',
        sachkonto_colname:str='Sachkonto-Nr.',
        report_type:str='Summary',
        main_groups_list:list[str]=['Aktiv','Leasing','Factoring','Bankverbindl.','Giro','Spar'],
        sachkonto_details_for:list[str]=['Giro']
        ):
    '''
    this function will read in the report and summarize the values by MSR position to be inserted in the main file
    '''
    if len(used_cost_centers)==0:
        print('All cost centers were unchecked, no report possible...')
        return None
    
    if not gesamt_sheet_name:
        print('No KUKA-Gesamtabfrage sheet was selected')
        return None
    
    if not fg_sheet_name:
        print('No KUKA-Finanzgeschäftabfrage sheet was selected')
        return None
    #reading in the report file
    # we are using the sheet_name and header row variables to immediately get to the table
    rdf = report_xl.parse(gesamt_sheet_name,header_row-1) #rdf -> report dataframe -> KUKA daten
    
    fgdf = report_xl.parse(fg_sheet_name,header_row-1)
    
    efg = fgdf.loc[fgdf[cost_center_colname].isin(used_cost_centers),[cost_center_colname,'Erträge Finanzgeschäft']].set_index(cost_center_colname) #Erträge Finanzgeschäft

    #print(f'Raw shape: {rdf.shape}')
    #print(rdf.columns)
    
    if 'Produkt' not in rdf.columns:
        print('There is no Produkt column found in the file, cannot run the further script. Please review the input parameters.')
        return None
        
    #we remove the lines where the Produkt column is not a number:
    rdf.loc[:,'Produkt'] = pd.to_numeric(rdf['Produkt'],errors='coerce')
    rdf = rdf.dropna(subset=['Produkt'])
    
    #print(f'Cleaned shape: {rdf.shape}')
    
    #we set the MSR category based on the dictionary defined previously - if something cannot be found, set it to "Undefined"
    rdf['MSR Cat'] = rdf['Produkt'].apply(lambda x: current_pt[x] if x in current_pt.keys() else "Undefined")
    
    #we create the category -> value matrix based on the below columns to generate the Active, Leasing, Giro, and Spar groups' numbers
    #the grouping to cost centers is in addition if this column is in the report and the option was selected earlier
    grouping_cols = ['MSR Cat']

    #if there is information about the SAKO in the report, we add this to the grouping columns
    if sachkonto_colname in rdf.columns:
        rdf = add_sako_categories(rdf)
        grouping_cols.append('SAKO Cat')
        sachkonto_flag = True
    else:
        sachkonto_flag = False
        
    if cost_center_colname in rdf.columns:
        imx_input = rdf.loc[rdf[cost_center_colname].isin(used_cost_centers)]
        if report_type=='Detailed':
            if len(used_cost_centers)>1:
                grouping_cols.append("Kostenstelle des Geschäfts")
            else:
                print('There is only one cost center selected, continuing with Summary.')
                report_type='Summary'      
    else:
        imx_input = rdf
        if report_type=='Detailed':
            print('Error - cannot found "Kostenstelle des Geschäfts" column in the report, cannot get details. Continuing with summary!')
            report_type='Summary'            

    
    imx = imx_input.groupby(grouping_cols, dropna=False).sum().loc[:,['Ø Saldo Aktiv','Ø Saldo Passiv',   #imx -> intermediate matrix
                                                'Ultimosaldo Aktiv','Ultimosaldo Passiv',
                                                'Zinsbeitrag Aktiv','Zinsbeitrag Passiv',
                                                'WP-Kurswerte']]#.map(lambda x: abs(x))
    
    #for each category, there is either passiv or active "average saldo", "ultimative saldo" and "zinsen" and we will extract these with a custom function
    first_part = get_first_part(imx_input=imx,
                                used_cost_centers = used_cost_centers,
                                group_list = main_groups_list,
                                sako_details_list = sachkonto_details_for,
                                detailed_flag = report_type=='Detailed',
                                sachkonto_flag = sachkonto_flag,
                                )

    #we use a function to convert the multiindex to simple string indices based on some rules to better match the insertion format
    first_part.index = [renaming_function(ind) for ind in first_part.index.tolist()]

    if report_type=='Detailed':
        second_part = get_second_part_detailed(
            imx_input,
            cost_center_colname,
            used_cost_centers
        )
    elif report_type=='Summary':
        second_part = get_second_part_summary(imx_input)

    print(f'Data types of parts matching: {type(first_part) is type(second_part)}')

    total=pd.concat([first_part,second_part])

    if isinstance(total,pd.Series):
        if len(used_cost_centers)>1:
            total.name='Summary'
        else:
            total.name=used_cost_centers[0]
        total = pd.concat([total[:-2],pd.Series(data=efg.sum(), index=['Erträge Finanzgeschäft'],name=total.name),total[-2:]])
    else:
        total = pd.concat([total[:-2], efg.T,total[-2:]])
    #hardcoding of making sure that only in case of report 230 the giro details are shown - shall be made a parameter in the future
    if used_cost_centers!=[230]:
        total.iloc[16:32]=np.nan
    return total

#details of getting the first part
sako_detail_map = [
    ('Girokonto',[1501,1502,1503]),
    ('Extrakonto',[2149,2250]),
    ('Pluskonto',[9191,9631,9671]),
    ('Fixzinskonto',[2143,2144,2251])
]

# ---------------------------------------------------------------------
# Creating a block in the first part of the report, based on the group_name
# ---------------------------------------------------------------------
def add_the_right_part(
        subser_input:pd.DataFrame,
        detailed_flag:bool,
        index_order_search_strings:list[str],
        cost_center_colname:str='Kostenstelle des Geschäfts',
        )->pd.Series|pd.DataFrame:
    
    #print('Raw input is: ',subser_input)
    
    #print("Raw input's index is: ",subser_input.index)
    subser = subser_input[subser_input.apply(lambda x:x!=0)].dropna()
    
    
    #print('subseries is: ', subser)
    #print('subseries index is: ', subser.index)
    
    #this is setting the columns' order
    old_order = subser.index.get_level_values('Position Types').tolist()
    order = setting_new_order(old_order,index_order_search_strings[:-1])  #we are trying to set the new index order (Except the % zins as this will be added later)
    #print(f'the new index order {order}')
    
    #we will reindex based on the new order and add the subdfs to a list (if an exact index is missing, we create a "row")
    extra_rows = []
    for pos in order:
        if pos not in subser.index.get_level_values(level='Position Types'):
            if detailed_flag:
                index = pd.MultiIndex.from_product([
                    subser_input.index.get_level_values('MSR Cat').unique(),
                    [pos],
                    subser_input.index.get_level_values(cost_center_colname).unique()
                    ],names=['MSR Cat', 'Position Types', cost_center_colname]
                )
            else:
                index = pd.MultiIndex.from_product([
                    subser_input.index.get_level_values('MSR Cat').unique(),
                    [pos],
                    ],names=['MSR Cat', 'Position Types']
                )
            extra_rows.append(pd.Series(data=[0]*len(index),index=index))

    subser = pd.concat([subser]+extra_rows).reindex(order, level='Position Types')
    #print(subser)
    #if this group is missing from the data, we still want to populate with np.NaNs 
    #for this it doesn't matter if we take the Active or Passive...
    if subser.empty:
        part = add_an_empty_part(category_name=subser_input.index.get_level_values('MSR Cat').unique()[0],
                      detailed_flag=detailed_flag,
                      cost_center_colname=cost_center_colname,
                      used_cost_centers=subser_input.index.get_level_values(cost_center_colname).unique()
        )
        return part
        
    if detailed_flag:
        top_part = subser.unstack(cost_center_colname)
        zins = pd.DataFrame(
            data = top_part.iloc[-1]/top_part.iloc[0],
            columns=[(top_part.index.get_level_values(0)[0],'Zinsbeitrag in %')]
            ).T
    else:
        top_part = subser
        zins = pd.Series(
            data = top_part.iloc[-1]/top_part.iloc[0],
            index=[(top_part.index.get_level_values(0)[0],'Zinsbeitrag in %')])
        
    part = pd.concat([top_part,zins],axis=0)
    
    #print('part added:')
    #print(part)
    return part

def add_an_empty_part(category_name:str,
                      detailed_flag:bool,
                      used_cost_centers:list[int]|None=None,
                      cost_center_colname:str|None=None,
                      pos_list:list[str]=['Ø Saldo Aktiv',  'Ultimosaldo Aktiv',  'Zinsbeitrag Aktiv','Zinsbeitrag in %']
                      ):
    if detailed_flag:
        if not cost_center_colname or not used_cost_centers:
            return None
        ser_index = pd.MultiIndex.from_product([
            [category_name],
            pos_list,
            used_cost_centers
            ],names=['MSR Cat', 'Position Types', cost_center_colname]
        )
    else:
        ser_index = pd.MultiIndex.from_product([
            [category_name],
            pos_list,
            ],names=['MSR Cat', 'Position Types']
        )
    
    subser = pd.Series(
            data=[0]*len(ser_index),
            index = ser_index
            )
    if detailed_flag:
        part = subser.unstack()
    else:
        part = subser
    
    return part


# ---------------------------------------------------------------------
# sorting the blocks (including the Sachkonto details)
# ---------------------------------------------------------------------
def get_sorting_order(group_list_input:list,sako_list:list,tpl_list:list=sako_detail_map)->list:
    group_list = group_list_input.copy()
    for val in sako_list:
        loc = group_list.index(val)
        for tpl in tpl_list[::-1]:
            group_list.insert(loc,f'{val} {tpl[0]}')
    return group_list

# ---------------------------------------------------------------------
# creating the first part of the report where per group always the Du. Saldo, the Ultimosaldo and the Zinsbeitrag is generated
# ---------------------------------------------------------------------
def get_first_part(imx_input:pd.DataFrame,
                   used_cost_centers:list[int|str],
                   group_list:list[str]=['Aktiv','Leasing','Giro','Spar'],
                   index_order_search_strings:list[str]=['Ø','Ultimosaldo','Zinsbeitrag','%'],
                   sako_details_list:list=['Giro'], # by default only apply the detailed SAKO breakdown to the Giro category
                   detailed_flag:bool=False,
                   sachkonto_flag:bool=True,
                   cost_center_colname:str='Kostenstelle des Geschäfts',
                   
                  )->pd.DataFrame|pd.Series:
    
    '''this summary function creates the first part of the insertable table/series by performing a series of reshaping, grouping, etc
    inputs: 
        imx: the matrix that is coming from the previous step (basically a bit reformed report): a dataframe with the MSR category and 
        cost center as multiindex and the report columns as columns
        group_list: the standard list of report categories, where the same operations need to be performed (default is given)
        index_order_search_strings: a list of strings that helps to align the index within each categories to match the order of the insertion
        sako_details_list: to which category (in the group list) shall a SAKO cat breakdown happen
        detailed_flag: the flag to decide whether to reuturn a summary as pd.Series or a detailed view via pd.Dataframe (with cost centers as columns)
    '''
    
    imx = imx_input.stack(future_stack=True)
    imx.index.set_names(imx.index.names[:-1]+['Position Types'],inplace=True)
         
    #sense_check
    if sachkonto_flag:
        sako2pop = []
        for i, sd_cat in enumerate(sako_details_list):
            if sd_cat not in group_list:
                sako2pop.append(i)
                print(f'Error! SAKO category "{sd_cat}" cannot be found in the report categories ({",".join(group_list)}), this group will be set to 0')
        if len(sako2pop)>0:
            for i in sako2pop[::-1]:
                sako_details_list.pop(i)
    else:
        sako_details_list = []
    
    subdfs = []
    #we only focus on the specific categories from the group list, if those combinations
    #do not have value (e.g. Factoring is missing), we still want to produce an empty dataframe
    for group in group_list:
        #print('Grouping value:',group)
        if group in imx.index.get_level_values('MSR Cat'):
            # as there are both active and passive categories of saldo, etc., we map which are bigger then 0
            # with the apply function, we keep those that are bigger than 0, because the rest will be np.NaN and will be removed by dropna()
            # (e.g. passive saldo row will be removed in case of the active group, because the whole row will be np.NaN.
            # This will leave us with 4 rows always (average saldo, ultimo saldo, zinsbeitrag EUR, zinsbeitrag %) 
            # it is possible though that some of the individual cost centers do not all have all the 4 values, so we need to use fillna(0) to show 0s instead
            # of empty cells
            this_group = imx.loc[[group]]
            if group in sako_details_list:
                #print(f'{group} break down by SAKO')
                #comes the SAKO breakdown - we need to do it by keys in the sako_detailed_map
                for tpl in sako_detail_map:
                    sako_cat=tpl[0]
                    if sako_cat not in this_group.index.get_level_values('SAKO Cat'):
                        print(f'creating empty data for "{group} {sako_cat}"')
                        #if this subcategory is completely missing, we just create an empty df
                        part = add_an_empty_part(category_name=f'{group} {sako_cat}',
                                                 detailed_flag=detailed_flag,
                                                 used_cost_centers=used_cost_centers,
                                                 cost_center_colname=cost_center_colname
                                                 )
                        subdfs.append(part)
                        continue

                    subgrp = this_group.loc[pd.IndexSlice[:,sako_cat,:],]                        
                    if detailed_flag:
                        grouping_levels = ['MSR Cat','Position Types',cost_center_colname]
                    else:
                        grouping_levels = ['MSR Cat','Position Types']
                    
                    subgrp_input = subgrp.groupby(level=grouping_levels).sum()
                    subgrp_input.index = subgrp_input.index.set_levels([f'{group} {sako_cat}'], level=subgrp_input.index.names[0])
                    
                    part = add_the_right_part(
                        subgrp_input,
                        detailed_flag,
                        index_order_search_strings
                    )
                    subdfs.append(part)

            if detailed_flag:
                
                part = add_the_right_part(
                    this_group.groupby(
                        level=[
                            'MSR Cat',
                            'Position Types',
                            cost_center_colname
                            ]
                        ).sum(),
                        detailed_flag,
                        index_order_search_strings
                )
                subdfs.append(part)
            else:
                part = add_the_right_part(
                    this_group.groupby(
                        level=[
                            'MSR Cat',
                            'Position Types'
                            ]
                        ).sum(),
                        detailed_flag,
                        index_order_search_strings
                )
                subdfs.append(part)
        else:
            #if there is no data, we add an empty dataframe/series
            part = add_an_empty_part(category_name=group,
                                     detailed_flag=detailed_flag,
                                     used_cost_centers=used_cost_centers,
                                     cost_center_colname=cost_center_colname
                                     )
            subdfs.append(part)
            
    #we get the first part by concatenating the selected dfs and filling np.NaNs with 0s and return it in the group_list's order
    result = pd.concat(subdfs).fillna(0)
    #print(result)
    sorting_order = get_sorting_order(group_list,sako_details_list)

    return result.reindex(sorting_order,level=0,axis=0)

# ---------------------------------------------------------------------
# function to add the SAKO categories, to the input dataframe based on the SAKO numbers in the sako_detail_map (list of tuples, not dict!)
# ---------------------------------------------------------------------
def add_sako_categories(imx:pd.DataFrame, sako_map:list=sako_detail_map, sachkonto_colname:str='Sachkonto-Nr.'):
    for tpl in sako_map:
        name=tpl[0]
        numbers=tpl[1]
        imx.loc[imx[sachkonto_colname].isin(numbers),'SAKO Cat']=name
    return imx


# ---------------------------------------------------------------------
# function to add the SAKO categories, to the input dataframe based on the SAKO numbers in the sako_detail_map (list of tuples, not dict!)
# ---------------------------------------------------------------------

def setting_new_order(current_order:list[str],search_strings:list[str])->list[str]:
    order=[]
        # the subframes all have 4 rows, but they are not in the right order for the report. we will reorder them by building a list of exact index names 
        # by using some keyword searches using regex this will allow us to find these regardless if they are passiv or active groups
    for s in search_strings:
        #print(f'searcing for {s} in index')
        for ind in current_order:
            #print(ind)
            m = re.search(s,ind)
            if m:
                #print(f'found it in this index: {ind}')
                order.append(ind)
                break
        else:
            order.append(s)

    return order


# ---------------------------------------------------------------------
# function to rename the blocks based on some rules to better align with the MSR target fields
# --------------------------------------------------------------------- 
def renaming_function(index_elements:list|tuple):
    '''this is a function te specifically rename the original index values for better alignment with the report's position. The rules below are 
    to produce optimal results for the current naming conventions, if something changes there, this code needs to be adjusted
    input: a (multi)index represented as an iterable (e.g.: list or tuple)
    '''
    cat = index_elements[0]
    #print('Category:',cat)
    if cat in ['Aktiv','Leasing','Factoring']:
        cat_out = cat+'volumen'
    elif cat in ['Giro','Spar']:
        cat_out = cat+'einlagen'
    elif cat=='Bankverbindl.':
        cat_out = 'Su.Bankverb.'
    else:
        cat_out = cat
    cost_type = index_elements[1]
    if not re.search('Zinsbeitrag', cost_type):
        if re.search('Ultimosaldo',cost_type):
            return cat_out +' - Saldo per Stichtag (in EUR)'
        else:
            return cat_out +' - Du.Stand (in EUR)'
    else:
        if re.search('%', cost_type):
            return 'Zinsbeitrag '+cat+' (in %)'
        else:
            return 'Zinsbeitrag '+cat+' (in EUR)'

# ---------------------------------------------------------------------
# creating the second part of the report where no passiv/aktiv columns need to be considered - this is for summary report
# ---------------------------------------------------------------------
def get_second_part_summary(imx_input:pd.DataFrame):
    second_part = imx_input.reindex([
    'Mindestreservekosten',
    'Zusatzerträge Aktiv',
    'Zusatzaufwände Aktiv',
    'Nicht ausgenützter Rahmen',
    'Risikokosten',
    'Eigenmittelkosten',
    'Liquiditätskosten Rahmen',
    'Liquiditätskosten Haftungen/Promessen',
    'Bonifikation Kreditsaldo',
    'Bonifikation Rahmen',
    'Wertstellungsnutzen',
    'Risikokosten nicht ausgenützte Rahmen',
    'Eigenmittelkosten nicht ausgenützte Rahmen',
    'Sonstige Erträge',
    'WP-Kurswerte',
    ],axis=1).sum()

    # this is to ensure that the "special" parts will be inserted in the right order
    first_split = 3 
    second_split = 11

    #inserting "Erträge aus Factoring"
    try:
        eaf = imx_input.loc[imx_input['MSR Cat']=='Factoring'].agg({'Erträge aus Factoring':'sum'})
    except KeyError:
        print('Error! "Erträge aus Factoring" column is missing from the report')
        eaf = pd.Series(data=[0],index=['Erträge aus Factoring'])
    
    #inserting "Ø Haftungsvolumen (EUR)"
    hv = imx_input.loc[(imx_input['MSR Cat']=='Haftungen')].agg({'Ø Saldo Haftung':'sum'}).set_axis(['Ø Haftungsvolumen (EUR)'])
    
    ewb = imx_input.loc[(imx_input['MSR Cat']=='EWB')].agg({'Ø Saldo Passiv':'sum'}).set_axis(['Einzelwertberichtigung (EUR)'])
    
    haft_acc = imx_input.groupby(['MSR Cat'],dropna=False).agg({'Ø Saldo Haftung':'sum'}).reindex([
        'Haftungen Bank',
        'Akkreditive',
        'Akkreditive Bank', # -> this does not have any product nummer associated with it yet
        ]).fillna(0).set_axis([
        'Haftungen Banken - Du.Stand (EUR)',
        'Akkreditive - Du.Stand (EUR)',
        'Akkreditive Banken - Du.Stand (EUR)',# -> this does not have any product nummer associated with it yet
        ]).squeeze()
    
    return pd.concat([second_part[:first_split], eaf, hv, ewb, second_part[first_split:second_split], haft_acc, second_part[second_split:]])

# ---------------------------------------------------------------------
# creating the second part of the report where no passiv/aktiv columns need to be considered - this is for detailed report
# ---------------------------------------------------------------------
def get_second_part_detailed(imx_input:pd.DataFrame, cost_center_colname:str,used_cost_centers:list[int]):
    second_part = imx_input.groupby([cost_center_colname]).sum().loc[:,[
    'Mindestreservekosten',
    'Zusatzerträge Aktiv',
    'Zusatzaufwände Aktiv',
    'Nicht ausgenützter Rahmen',
    'Risikokosten',
    'Eigenmittelkosten',
    'Liquiditätskosten Rahmen',
    'Liquiditätskosten Haftungen/Promessen',
    'Bonifikation Kreditsaldo',
    'Bonifikation Rahmen',
    'Wertstellungsnutzen',
    'Risikokosten nicht ausgenützte Rahmen',
    'Eigenmittelkosten nicht ausgenützte Rahmen',
    'Sonstige Erträge',
    'WP-Kurswerte',
    ]].reindex(used_cost_centers)
    
    #inserting "Erträge aus Factoring"
    eaf = imx_input.loc[imx_input['MSR Cat']=='Factoring'].groupby([cost_center_colname]).sum().loc[:,[
        'Erträge aus Factoring']].squeeze().reindex(used_cost_centers)
    
    #inserting "Ø Haftungsvolumen (EUR)"
    hv = imx_input.loc[(imx_input['MSR Cat']=='Haftungen')
                    ].groupby([cost_center_colname
                               ],dropna=False).agg({'Ø Saldo Haftung':'sum'}).set_axis(['Ø Haftungsvolumen (EUR)'],axis=1).squeeze().reindex(used_cost_centers)
    
    ewb = imx_input.loc[(imx_input['MSR Cat'
                         ]=='EWB')].groupby([cost_center_colname
                                             ],dropna=False).agg({'Ø Saldo Passiv':'sum'}).set_axis(['Einzelwertberichtigung (EUR)'
                                                                                                        ],axis=1).squeeze().reindex(used_cost_centers)
    
    haft_acc = imx_input.groupby([cost_center_colname,
                       'MSR Cat'],dropna=False).sum().loc[:,'Ø Saldo Haftung'].unstack(level=1).reindex([
        'Haftungen Bank',
        'Akkreditive',
        'Akkreditive Bank', # -> this does not have any product nummer associated with it yet
        ], axis=1).fillna(0).set_axis([
        'Haftungen Banken - Du.Stand (EUR)',
        'Akkreditive - Du.Stand (EUR)',
        'Akkreditive Banken - Du.Stand (EUR)',# -> this does not have any product nummer associated with it yet
        ],axis=1).reindex(used_cost_centers)
    
    #because this is a dataframe, we can insert the "columns"
    locations = [3,4,5]
    
    entries = [eaf,hv,ewb]
    for i, entry in enumerate(entries):
        second_part.insert(locations[i],entry.name, entry.values)
    
    #because this is a dataframe, we can insert the "columns"
    second_locations = [14,15,16]
    for i, col in enumerate(haft_acc.columns):
        second_part.insert(second_locations[i], col, haft_acc[col])

    return second_part.T.fillna(0)


def euro_format(x, decimals=2):
    """
    Formats a numeric value using '.' for thousands
    and ',' for decimal separator, to 'decimals' places.
    Example: 1234567.89 -> "1.234.567,89"
    """
    if pd.isna(x):
        return ""
    # First format with standard (U.S.) grouping: e.g. "1,234,567.89"
    # The spec ":,.2f" means comma-grouping and 2 decimal places.
    us_style = f"{x:,.{decimals}f}"

    # Swap commas (,) and periods (.) to get the European style:
    # 1) Temporarily replace commas with a marker ("X")
    # 2) Replace periods with commas
    # 3) Replace "X" with periods
    # E.g. "1,234,567.89" -> "1X234X567,89" -> "1X234X567,89" -> "1.234.567,89"
    euro_style = us_style.replace(",", "X").replace(".", ",").replace("X", ".")
    return euro_style

def custom_formatter(val,decimals:int=2):
    """
    If abs(x) >= 1, return it with underscore thousands separators, 0 decimals.
    If abs(x) < 1, interpret as a fraction (0.02 -> "2.00%").
    Adjust logic as needed for negative values or special cases.
    """
    if pd.isna(val):
        return ""
    elif val==0:
        return f"{int(val)}"
    elif abs(val) < 1:
        # Show as a percentage with 2 decimals
        return f"{val:.{decimals}%}"
    else:
        return euro_format(val, decimals)

def update_target_excel_xlwings(
    data_input:pd.Series|pd.DataFrame,
    cost_centers_used:list[int],
    target_column:str,
    date_row:int=8,
    data_row:int=115,
    titlestr:str|None=None
    ):
    
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    target_path = filedialog.askopenfilename(title=titlestr, filetypes=[('Excel files','.xlsx')])
    
    if not target_path:
        print('No target file was selected!')
        root.destroy()    
        return
    app = xw.App(add_book=False)

    wb=app.books.open(target_path,update_links=False)
    root.destroy()
    for cc in cost_centers_used:
        print(f'updating cost_center {cc}')
        if isinstance(data_input,pd.Series): #this is in case of summary report
            data_to_insert = data_input
        elif isinstance(data_input,pd.DataFrame):
            data_to_insert = data_input[cc]

        for sheet in wb.sheet_names:
            if str(cc) in sheet:
                ws = wb.sheets[sheet]
                print(f'Found the worksheet for {str(cc)}')
                break #Found the sheet to paste the data into!
        else:
            print(f"Did not find any sheet with {str(cc)} in it's name, no update")
            ws = None
            continue

        for i, cell in enumerate(ws[f'{date_row}:{date_row}']): #the rows where the date is defined for the column
            #print(i)
            if i>=150:
                print(f"Did not find {target_column} in the first 150 columns of row {date_row} in worksheet {sheet}, no update")            
                target_col = None
                break

            if cell.value==target_column:
                target_col = cell.column
                address = cell.get_address()
                print(f'Found the column to update for worksheet {sheet}:{address}')
                break #Found the column to paste the data into!
        else:
            print(f"Did not find any column with {target_column} in row {date_row} in worksheet {sheet}, no update")            
            target_col = None
            continue

        if target_col:
            print(f'pasting data into {ws.range(data_row,target_col).get_address()}')
            ws.range(data_row,target_col).value = [[val] for val in data_to_insert.values.tolist()]
        else:
            continue

    print('Update completed, saving & closing workbook')
    wb.save(target_path)
    wb.close()
    app.quit()

def main():
    pass
    #test_xlwings(pd.Series([1212,3232,43535,67888,34343],index=['haha','hehe','hihi','lolo','püéö'],name=230),
    #             [230],
    #             '12/2024',
    #)

if __name__=='__main__':
    main()
