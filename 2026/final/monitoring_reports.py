# -*- coding: utf-8 -*-
"""
This scripts generates the cluster-briefs.

Author: Jeff Slee
Date Created: May 2023
Last updated: 11 August 2026

"""
#%% Read libraries
# standard libraries
import os

import re
import unicodedata
import pandas as pd
import numpy as np
import math

# # import graphics
import matplotlib.pyplot as plt
import seaborn as sns

# # import fonts
from matplotlib.font_manager import FontProperties
import matplotlib.ticker as mticker

# import docx
from docxtpl import DocxTemplate, InlineImage
from datetime import datetime
from docx.shared import Mm
# from num2words import num2words
            
#%% Setting file paths

## Set GII years
giiyr = 2026
giiyr = str(giiyr)

# change working directory
os.chdir('C:/python/monitoring_reports/' + str(giiyr))


# from decimal import Decimal, ROUND_UP
pd.options.display.float_format = '{:,}'.format

curtime = datetime.today().strftime('%m%d%y_%H%M')

## Files
indicator_status_file   = './data/indicator_status_2026.csv'
indicator_counts_file   = './data/indicator_counts_2026.csv'
templatePath            = './template/template.docx'

### for ranks
SUFFIXES = {1: 'ˢᵗ', 2: 'ⁿᵈ', 3: 'ʳᵈ'}
def ordinal(num):
    # I'm checking for 10-20 because those are the digits that
    # don't follow the normal counting scheme. 
    if 10 <= num % 100 <= 20:
        suffix = 'ᵗʰ'
    else:
        # the second parameter is a default.
        suffix = SUFFIXES.get(num % 10, 'ᵗʰ')
    return str(num) + suffix

def adjname(name):
    s = unicodedata.normalize('NFKD', name)
    s = s.encode('ascii', 'ignore').decode('ascii')   # drop accents
    s = re.sub(r'[^A-Za-z0-9]+', '_', s)              # anything else -> _
    return s.strip('_').lower()

#%% Reading in data

## Reading data files
df_indicator_counts_full = pd.read_csv(indicator_counts_file, header=0)
df_indicator_counts_full['ECONOMY_NAME_CLN'] = (
                                                df_indicator_counts_full['ECONOMY_NAME'].map(adjname)
                                            )

### setting nans to 0 - add to prep script later
p_cols = ['P1', 'P2', 'P3', 'P4', 'P5', 'P6', 'P7']
df_indicator_counts_full[p_cols] = df_indicator_counts_full[p_cols].fillna(0).astype(int)

df_indicator_status_full = pd.read_csv(indicator_status_file, header=0).query("STATUS in ['MISSING', 'OUTDATED']")\
                            .reset_index(drop=True)\
                                .assign(DATAYR=lambda d: d["DATAYR"].astype("Int64").astype("string").fillna(""),
                                        MODEYR=lambda d: d["MODEYR"].astype("Int64").astype("string").fillna(""))

max_input_indicators = df_indicator_counts_full["IN_INDICATORS"].drop_duplicates().iloc[0]
max_output_indicators = df_indicator_counts_full["OUT_INDICATORS"].drop_duplicates().iloc[0]-1 ## subtract utility models

input_dmc = math.floor(max_input_indicators * .66)
output_dmc = math.floor(max_output_indicators * .66)

imax = max_input_indicators - input_dmc - 1 ## !not sure why Lorena says 18!
omax = max_output_indicators - output_dmc

#%% Getting cluster specific dataframe
all_economies = df_indicator_counts_full["ISO3"].drop_duplicates().tolist()
nongii_economies = df_indicator_counts_full[(df_indicator_counts_full['GIIECON'] != 1) | 
                                            (df_indicator_counts_full['GIIECON'].isna())]["ISO3"].drop_duplicates().tolist()
economies = nongii_economies 

### regional aggregates
df_region = df_indicator_counts_full[['ECONOMY_NAME', 'ISO3', 'S_MISSING', 
                                      'S_OUTDATED', 'SUBREG_UN', 'INTREG_UN',
                                      'GIIECON']].copy()

## inter-regional ranks - rank within each region; 1 = lowest value
grp = df_region.groupby('INTREG_UN')
df_region['S_MISSING_RANK'] = grp['S_MISSING'].rank(
    method='min', ascending=True
).astype(int)

df_region['S_OUTDATED_RANK'] = grp['S_OUTDATED'].rank(
    method='min', ascending=True
).astype(int)


def region_stats(df, group_col, share_col=None):
    
    if share_col is None:
        raise ValueError("Column share - S_MISSING or S_OUTDATED is required")
        
    g = df.groupby(group_col)[share_col]
    out = pd.DataFrame({
        "n":    g.size(),
        "min":  g.min(),
        "q1":   g.quantile(.25),
        "q2":   g.median(),
        "q3":   g.quantile(.75),
        "max":  g.max(),
        "mean": g.mean(),
    })

    # # economy holding the extreme value in each region
    # def holder(idx, label):
    #     h = df.loc[idx, [group_col, "ECONOMY_NAME", share_col, "GIIECON"]].copy()
    #     h["GIIECON"] = h["GIIECON"].eq(1).map({True: "GII", False: "Non-GII"})
    #     return (h.set_index(group_col)
    #              .rename(columns={"ECONOMY_NAME": label + "_econ",
    #                               share_col:      label + "_val",
    #                               "GIIECON":      label + "_gii"}))

    # out = out.join(holder(g.idxmin(), "least_missing"))     # lowest missing count
    # out = out.join(holder(g.idxmax(), "highest_missing"))   # highest missing count
    return out.round(1)

#### Inter-Regional stats
intreg_stats_missing = region_stats(df_region, "INTREG_UN", 'S_MISSING').reset_index()
intreg_stats_outdated = region_stats(df_region, "INTREG_UN", 'S_OUTDATED').reset_index()
    
# # Set colors for the plots
# colours = ['#EC651F','#76B82A','#23B9D6','#8C96B1',
#            '#EAB494','#C2D99C','#A3DAE8','#CED2DE']

#%% Plot settings

sns.set()
sns.set_style("white")

os.makedirs('./graphs', exist_ok=True)

# Set colors for the plots
colours = ['#EC651F','#76B82A','#23B9D6','#8C96B1',
           '#EAB494','#C2D99C','#A3DAE8','#CED2DE']

C_FOCUS    = colours[0]   # the reporting economy       orange
C_MED      = colours[3]   # medians / reference lines   grey-blue
C_REGION   = colours[6]   # rest of the sub-region      light blue
C_WORLD    = colours[7]   # all other economies         pale grey
C_MISSING  = '#FF9FAA'    # missing indicators          (255,159,170)
C_OUTDATED = '#23B9D6'    # outdated indicators         (35,185,214)

title_font = FontProperties(
    fname=r'C:\Users\Jeffrey Slee\AppData\Local\Microsoft\Windows\Fonts\NotoSansDisplay-Bold.ttf',
    size=22
)
label_font = FontProperties(family="Noto Sans Display", size=15)

# Target dimensions in cm
box_width_cm = 16.5
box_height_cm = 12.0

cm_to_in = 1 / 2.54
scale = 3   # render 3x the physical target, Word scales it back down

# medians across all economies, used as the reference in both plots
med_mis = df_indicator_counts_full["S_MISSING"].median()
med_out = df_indicator_counts_full["S_OUTDATED"].median()

# medians across GII economies only
df_gii = df_indicator_counts_full[df_indicator_counts_full["GIIECON"] == 1]
gii_med_mis = df_gii["S_MISSING"].median()
gii_med_out = df_gii["S_OUTDATED"].median()

C_GII = '#76B82A'

# # official names too long for the bar-chart gutter
# label_short = {'Iran (Islamic Republic of)': 'Iran (Isl. Rep. of)',
#                'Venezuela (Bolivarian Republic of)': 'Venezuela (Bol. Rep. of)',
#                'Democratic Republic of the Congo': 'DR Congo',
#                'United Republic of Tanzania': 'Tanzania',
#                'Bolivia (Plurinational State of)': 'Bolivia',
#                "Lao People's Democratic Republic": 'Lao PDR',
#                'Micronesia (Federated States of)': 'Micronesia',
#                'Saint Vincent and the Grenadines': 'St Vincent & Gren.'}

rename_cols = {
                "index": "N",
                "SUBINDEX_NAME": "SN",
                "PILLAR_NAME": "PN",
                "SUBPILLAR": "SC",
                "INDICATOR": "IC",
                "DATAYR": "DY",
                "MODEYR": "MY",
                "SOURCE": "S",
            }    


###### START LOOP FOR EACH CLUSTER HERE
## Temp override, REMOVE LATER!!!!!!!!!
# iso3 = economies[1]
# iso3 = economies[2]
# iso3 = economies[18]
# iso3 = economies[98]
for iso3 in economies: #all_clusters: #[0:3]:
    
    print(iso3)
    
    ######------------------------------------------------------------------
    # filter for individual economy data
    df_indicator_counts_ind = df_indicator_counts_full.loc[ 
                              df_indicator_counts_full["ISO3"] == iso3].\
                                    reset_index(drop=True)                                                                          
                                    
    df_indicator_status_missing_ind = (
                                        df_indicator_status_full
                                        .query("ISO3 == @iso3 and STATUS in ['MISSING']")
                                        .assign(index=lambda d: range(1, len(d) + 1))
                                        .rename(columns=rename_cols)
                                        [list(rename_cols.values())]
                                    )
                                                                                    
                                            
    df_indicator_status_outdated_ind = (
                                        df_indicator_status_full
                                        .query("ISO3 == @iso3 and STATUS in ['OUTDATED']")
                                        .assign(index=lambda d: range(1, len(d) + 1))
                                        .rename(columns=rename_cols)
                                        [list(rename_cols.values())]
                                    )                                           
                                    
    ##### --- benchmarking
    # econ_subregion = df_indicator_counts_ind.at[0,"SUBREG_UN"]
    econ_intregion = df_indicator_counts_ind.at[0,"INTREG_UN"]
    

    df_bench = df_indicator_counts_full.loc[ 
                df_indicator_counts_full["INTREG_UN"] == econ_intregion].\
                  reset_index(drop=True).rename(columns={"ECONOMY_NAME": "LABEL"})
   
    #### region's rank
    df_region_rank_ind = (
                            df_region
                            .query("ISO3 == @iso3")
                            [['ISO3', 'S_MISSING_RANK', 'S_OUTDATED_RANK']]
                                .reset_index(drop=True)
                        )
                                                            
    
    #### regional stats
    intreg_stats_missing_ind = intreg_stats_missing.loc[ 
                                   intreg_stats_missing["INTREG_UN"] == econ_intregion].\
                                     reset_index(drop=True)  
                                     
    intreg_stats_outdated_ind = intreg_stats_outdated.loc[ 
                                   intreg_stats_outdated["INTREG_UN"] == econ_intregion].\
                                     reset_index(drop=True)                                       
                                    
#%% Create plots

    ## Plot 1 - the sub-region ranked on missing and on outdated shares
    fig, axes = plt.subplots(1, 2,
                             figsize=(box_width_cm * cm_to_in * scale, box_height_cm * cm_to_in * scale),
    constrained_layout=True)

    for ax, col, ttl, bar_col, med, gii_med in [
            (axes[0], "S_MISSING",  "Indicators missing (%)",  C_MISSING,  med_mis, gii_med_mis),
            (axes[1], "S_OUTDATED", "Indicators outdated (%)", C_OUTDATED, med_out, gii_med_out)]:
        
        d = df_bench.sort_values(col)
        ax.barh(d["LABEL"], d[col], height=0.72, zorder=3,
                color=[C_FOCUS if i == iso3 else bar_col for i in d["ISO3"]])
        
        # Put lowest value at the top and highest at the bottom
        ax.invert_yaxis()

        # xlim first, and wide enough to hold both medians
        ax.set_xlim(0, max(d[col].max() * 1.18, med * 1.18, gii_med * 1.18, 10))
        
        ax.axvline(med, color=C_MED, linestyle='--', linewidth=2, zorder=5)

        # Place label at top using axis transform (y = 0.98 places it near top)
        ax.annotate("Global median {:.0f}%".format(med),
                    xy=(med, 0.98), xycoords=ax.get_xaxis_transform(),
                    xytext=(4, 0), textcoords='offset points', 
                    color=C_MED, fontproperties=label_font, 
                    va='top', ha='left')

        ax.axvline(gii_med, color=C_GII, linestyle='--', linewidth=2, zorder=5)
        flip = gii_med > ax.get_xlim()[1] * 0.62      # keep the label inside the panel
        ax.annotate("GII median {:.0f}%".format(gii_med),
                    xy=(gii_med, 0.012), xycoords=ax.get_xaxis_transform(),
                    xytext=(-4 if flip else 4, 0), textcoords='offset points',
                    color=C_GII, fontproperties=label_font,
                    ha='right' if flip else 'left', va='bottom')

        for y, (v, i) in enumerate(zip(d[col], d["ISO3"])):
            ax.annotate("{:.0f}".format(v), xy=(v, y), xytext=(5, 0),
                        textcoords='offset points', va='center',
                        fontproperties=label_font,
                        color=C_FOCUS if i == iso3 else '#201E1D')

        ax.set_title(ttl, fontproperties=title_font, loc='left', pad=10)
        ax.set_xlim(0, max(d[col].max() * 1.18, 10))
        ax.xaxis.set_major_formatter(mticker.PercentFormatter(decimals=0))
        ax.grid(axis='x', color='#E4E4E6', zorder=0)
        ax.set_axisbelow(True)
        for side in ['top', 'right', 'left']:
            ax.spines[side].set_visible(False)
        ax.spines['bottom'].set_color(C_WORLD)
        ax.tick_params(length=0)
        for lbl in ax.get_yticklabels() + ax.get_xticklabels():
            lbl.set_fontproperties(label_font)

    b_filepath = './graphs/bench_' + iso3 + '.png'
    plt.savefig(b_filepath, dpi=300, format="png")
    plt.close(fig)
                                      
    
#%% Create variables for template
    
    ## set variables for report
    econ = str(df_indicator_counts_ind.at[0,"ECONOMY_NAME"])
    econ_cln = str(df_indicator_counts_ind.at[0,"ECONOMY_NAME_CLN"])
    giiyr = str(giiyr)
    idmc = str(input_dmc)
    odmc = str(output_dmc)
    imax = str(imax)
    omax = str(omax)
    
    r_tot   = int(df_indicator_counts_ind.at[0,"N_REPORTED"])
    r_in   = int(df_indicator_counts_ind.at[0,"IN_REPORTED"])
    r_out   = int(df_indicator_counts_ind.at[0,"OUT_REPORTED"])
    
    n_tot   = int(df_indicator_counts_ind.at[0,"N_NEW"])
    n_in   = int(df_indicator_counts_ind.at[0,"IN_NEW"])
    n_out   = int(df_indicator_counts_ind.at[0,"OUT_NEW"])
    
    o_tot   = int(df_indicator_counts_ind.at[0,"N_OUTDATED"])
    o_in   = int(df_indicator_counts_ind.at[0,"IN_OUTDATED"])
    o_out   = int(df_indicator_counts_ind.at[0,"OUT_OUTDATED"])
    
    o_tot_14y   = int(df_indicator_counts_ind.at[0,"N_OUTDATED_1_4"])
    o_in_14y   = int(df_indicator_counts_ind.at[0,"IN_OUTDATED_1_4"])
    o_out_14y   = int(df_indicator_counts_ind.at[0,"OUT_OUTDATED_1_4"])  
    
    o_tot_59y   = int(df_indicator_counts_ind.at[0,"N_OUTDATED_5_9"])
    o_in_59y   = int(df_indicator_counts_ind.at[0,"IN_OUTDATED_5_9"])
    o_out_59y   = int(df_indicator_counts_ind.at[0,"OUT_OUTDATED_5_9"])    
    
    o_tot_10y   = int(df_indicator_counts_ind.at[0,"N_OUTDATED_10"])
    o_in_10y   = int(df_indicator_counts_ind.at[0,"IN_OUTDATED_10"])
    o_out_10y   = int(df_indicator_counts_ind.at[0,"OUT_OUTDATED_10"])  
    
    m_tot   = int(df_indicator_counts_ind.at[0,"N_MISSING"])
    m_in   = int(df_indicator_counts_ind.at[0,"IN_MISSING"])
    m_out   = int(df_indicator_counts_ind.at[0,"OUT_MISSING"])  
    
    s_rep =  "{:.0%}".format((df_indicator_counts_ind.at[0,"S_REPORTED"]/100))
    s_new =  "{:.0%}".format((df_indicator_counts_ind.at[0,"S_NEW"]/100))
    s_out =  "{:.0%}".format((df_indicator_counts_ind.at[0,"S_OUTDATED"]/100))
    s_mis =  "{:.0%}".format((df_indicator_counts_ind.at[0,"S_MISSING"]/100))
    
    p1 =  int(df_indicator_counts_ind.at[0,"P1"])
    p2 =  int(df_indicator_counts_ind.at[0,"P2"])
    p3 =  int(df_indicator_counts_ind.at[0,"P3"])
    p4 =  int(df_indicator_counts_ind.at[0,"P4"])
    p5 =  int(df_indicator_counts_ind.at[0,"P5"])
    p6 =  int(df_indicator_counts_ind.at[0,"P6"])
    p7 =  int(df_indicator_counts_ind.at[0,"P7"])

    
    ### regional stats
    n = int(intreg_stats_missing_ind.at[0,"n"])
    
    ### missing
    m_min = "{:.1%}".format((intreg_stats_missing_ind.at[0,"min"]/100))
    m_q1 = "{:.1%}".format((intreg_stats_missing_ind.at[0,"q1"]/100))
    m_q2 = "{:.1%}".format((intreg_stats_missing_ind.at[0,"q2"]/100))
    m_q3 = "{:.1%}".format((intreg_stats_missing_ind.at[0,"q3"]/100))
    m_max = "{:.1%}".format((intreg_stats_missing_ind.at[0,"max"]/100))
    m_avg = "{:.1%}".format((intreg_stats_missing_ind.at[0,"mean"]/100))
    m_rank = ordinal(int(df_region_rank_ind.at[0,"S_MISSING_RANK"]))
    
    # m_econ_min_nm = str(intreg_stats_missing_ind.at[0,"least_missing_econ"]) + " economy"
    # m_econ_min_gii = str(intreg_stats_missing_ind.at[0,"least_missing_gii"])
    # m_econ_max_nm = str(intreg_stats_missing_ind.at[0,"highest_missing_econ"]) + " economy"
    # m_econ_max_gii = str(intreg_stats_missing_ind.at[0,"highest_missing_gii"])
    
    ### outdated
    o_min = "{:.1%}".format((intreg_stats_outdated_ind.at[0,"min"]/100))
    o_q1 = "{:.1%}".format((intreg_stats_outdated_ind.at[0,"q1"]/100))
    o_q2 = "{:.1%}".format((intreg_stats_outdated_ind.at[0,"q2"]/100))
    o_q3 = "{:.1%}".format((intreg_stats_outdated_ind.at[0,"q3"]/100))
    o_max = "{:.1%}".format((intreg_stats_outdated_ind.at[0,"max"]/100))
    o_avg = "{:.1%}".format((intreg_stats_outdated_ind.at[0,"mean"]/100))
    o_rank = ordinal(int(df_region_rank_ind.at[0,"S_OUTDATED_RANK"]))
   
    # o_econ_min_nm = str(intreg_stats_outdated_ind.at[0,"least_missing_econ"])
    # o_econ_min_gii = str(intreg_stats_outdated_ind.at[0,"least_missing_gii"]) + " economy"
    # o_econ_max_nm = str(intreg_stats_outdated_ind.at[0,"highest_missing_econ"])
    # o_econ_max_gii = str(intreg_stats_outdated_ind.at[0,"highest_missing_gii"]) + " economy"

    
    # Filename
    econfilename = "gii_" + giiyr + "_monitoring_report_" + econ_cln
    
#%% Generate report
    
    doc = DocxTemplate(templatePath)
       
    table_1 = []        
    for idx, row in df_indicator_status_missing_ind.iterrows():
        tmp = row.to_dict()
        table_1.append(tmp)
    
    table_2 = []        
    for idx, row in df_indicator_status_outdated_ind.iterrows():
        tmp = row.to_dict()
        table_2.append(tmp)
    
    ## Graphs
    g_bench = InlineImage(doc, b_filepath, width=Mm(165), height=Mm(120))
    
    # # ## Maps
    # g_map = InlineImage(doc,m_filepath, width=Mm(125), height=Mm(125))
    
    context = {
                'econ':  econ,
                'giiyr': giiyr,
                'idmc':  idmc,
                'odmc':  odmc,
                'imax':  imax,
                'omax':  omax,
                
                'r1': r_tot,
                'r2': r_in,
                'r3': r_out,

                'n1': n_tot,
                'n2': n_in,
                'n3': n_out,
                
                'o1': o_tot,
                'o2': o_in,
                'o3': o_out,
                
                'o1_14y': o_tot_14y,
                'o2_14y': o_in_14y,
                'o3_14y': o_out_14y,
                
                'o1_59y': o_tot_59y,
                'o2_59y': o_in_59y,
                'o3_59y': o_out_59y,
                
                'o1_10y': o_tot_10y,
                'o2_10y': o_in_10y,
                'o3_10y': o_out_10y,
                
                'm1': m_tot,
                'm2': m_in,
                'm3': m_out,
                
                's_rep': s_rep,
                's_new': s_new,
                's_out': s_out,
                's_mis': s_mis,
                
                'p1': p1,
                'p2': p2,
                'p3': p3,
                'p4': p4,
                'p5': p5,
                'p6': p6,
                'p7': p7,
                
                'intreg': econ_intregion,
                'n': n,
                
                'm_min': m_min,
                'm_q1': m_q1,
                'm_q2': m_q2,
                'm_q3': m_q3,
                'm_max': m_max,
                'm_avg': m_avg,
                'm_rank': m_rank,
                # 'm_econ_min_nm': m_econ_min_nm,
                # 'm_econ_min_gii': m_econ_min_gii,
                # 'm_econ_max_nm': m_econ_max_nm,
                # 'm_econ_max_gii': m_econ_max_gii,
                
                'o_min': o_min,
                'o_q1': o_q1,
                'o_q2': o_q2,
                'o_q3': o_q3,
                'o_max': o_max,
                'o_avg': o_avg,
                'o_rank': o_rank,
                # 'o_econ_min_nm': o_econ_min_nm,
                # 'o_econ_min_gii': o_econ_min_gii,
                # 'o_econ_max_nm': o_econ_max_nm,
                # 'o_econ_max_gii': o_econ_max_gii,
                
                'g_bench': g_bench,
                
                'table_1': table_1,
                'table_2': table_2,
                
               }
    
    
    doc.render(context, autoescape=True)
    

    docx_file = (econfilename + '.docx').lower()
    docx_path = os.path.abspath('./final/word/' + docx_file)
    doc.save(docx_path)
        
    
    
    # del clustfilename, rank_num, prev_rank_num, rank_chg, rank_percap_num, rank_percap_chg,\


