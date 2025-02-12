# -*- coding: utf-8 -*-
"""
HIGH SPEED COMMUNICATION NETWORKS LABORATORY
NATIONAL TECHNICAL UNIVERSITY OF ATHENS

Created on Tue Jan 21 18:27:27 2025

@author: Alireza Khaksari
@email:  alirezakfz@mail.ntua.gr
@mail:   alireza_kfz@yahoo.com
  

"""

import matplotlib.pyplot as plt
import seaborn as sn
import numpy as np
import pandas as pd
import os
import sys
import re
import openpyxl


file_name="All_Results.xlsx"

fig_path = os.getcwd()
file_path =  fig_path = os.getcwd()
file_path =  os.path.join(fig_path, file_name)

algorithms = ["Comp", "Diag", "FP", "SFP"]
DAs =['DA1', 'DA2', 'DA3', 'DA4']

if os.path.exists(file_path):
    print("File Exist")
else:
    print("Correct the path")
    
    
def plot_smps_comp_vs_learning(file_path, algorithms, fig_path):
    sheet_name = "LMPS"
    
    # Number of columns to plot 
    ncols = len(algorithms) -1 
    
    # Number of rows to plot
    nrows = 1
    
    fig, axes = plt.subplots(nrows , ncols, figsize=(ncols*12, 8))
    
    alphabet = 'a'
    
    # Select the first value as the pivot
    pivot = algorithms[0]
    
    row =0
    count = 0
    
    value_names = {'Comp':"Competitive",
                        "SFP":"Smooth Fictitious Play",
                        "FP":"Fictitious Play",
                        "Diag":"Diagonalization"}
    
    palette_colors = sn.color_palette('tab10')
    palette_dict = {model: color for model, color in zip(algorithms, palette_colors)}
    
   
    if(os.path.exists(file_path)):
        # Load the data 
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        for algo in algorithms[1:]:
            ax = axes[row]
            df_temp = df[df['Algorithm'].isin([pivot, algo])]
            sn.barplot(data= df_temp, y='Calculated_LMPs', x='MTU', hue='Algorithm', ax=ax, palette= palette_dict)
            
            # Title font size
            ax.set_title('({0})-{1} vs {2}'.format(alphabet, 'Competitive', value_names[algo]), fontsize=18)
             
            # X and Y label font sizes
            ax.set_xlabel('MTU (h)', fontsize=18)
            ax.set_ylabel('SMP(\u20AC/MWh)', fontsize=18)
            
            # Tick label font size
            ax.tick_params(axis='both', which='major', labelsize=18)
            
            # Legend font size
            ax.legend(fontsize=16)
            
            row+=1
            
            count += 1
            alphabet = chr(ord('a') + count)
       
        
        save_path = os.path.join(fig_path, 'SMPs_All_Results_Competitive_VS_Others.png')
        eps_path = os.path.join(fig_path, 'SMPs_All_Results_Competitive_VS_Others.eps')
        
        print(save_path)
        fig.savefig(save_path,  bbox_inches="tight")
        fig.savefig(eps_path,  bbox_inches="tight", format="eps", dpi=1200)
        
    
    
plot_smps_comp_vs_learning(file_path, algorithms, fig_path)


def plot_DAs_bids(file_path, algorithms,DAs ,fig_path):
    sheet_name = "BIDS_MVA"
    
    # Number of columns to plot 
    ncols = len(DAs)
    
    # Number of rows to plot
    nrows = 1
    
    # Select the first value as the pivot
    pivot = algorithms[0]
    
    value_names = {'Comp':"Competitive",
                        "SFP":"Smooth Fictitious Play",
                        "FP":"Fictitious Play",
                        "Diag":"Diagonalization"}
    
    palette_colors = sn.color_palette('tab10')
    palette_dict = {model: color for model, color in zip(algorithms, palette_colors)}
    
    if(os.path.exists(file_path)):
        # Load the data 
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    
        for algo in algorithms[1:]:
            
            alphabet = 'a'
            fig, axes = plt.subplots(nrows , ncols, figsize=(ncols*12, 8))
            
            row =0
            count = 0
            
            for da in DAs:
                ax = axes[row]
                df_temp = df[df['Algorithm'].isin([pivot, algo])]
                df_temp = df_temp[['MTU', da, 'Algorithm']].copy()
                
                sn.barplot(data= df_temp, y=da, x='MTU', hue='Algorithm', ax=ax, palette= palette_dict)
    
                # Title font size
                ax.set_title('({0})-{1} Bids'.format(alphabet, da), fontsize=18)
                 
                # X and Y label font sizes
                ax.set_xlabel('MTU (h)', fontsize=18)
                ax.set_ylabel('SMP(\u20AC/MWh)', fontsize=18)
                
                # Tick label font size
                ax.tick_params(axis='both', which='major', labelsize=18)
                
                # Legend font size
                ax.legend(fontsize=16)
                
                row+=1
                
                count += 1
                alphabet = chr(ord('a') + count)
                
            
            save_path = os.path.join(fig_path, 'DAs_Bids_Competitive_VS_{0}.png'.format(algo))
            eps_path = os.path.join(fig_path, 'DAs_Bids_Competitive_VS_{0}.eps'.format(algo))
            
            print(save_path)
            fig.savefig(save_path,  bbox_inches="tight")
            fig.savefig(eps_path,  bbox_inches="tight", format="eps", dpi=1200)
            

plot_DAs_bids(file_path, algorithms,DAs ,fig_path)




def plot_algorithms_convergence(convergence_file_path, algorithms, DAs ,fig_path):
    sheet_names = ["DIAG_PLOT", "FP_PLOT", "SFP_PLOT"]
    
    # Number of columns to plot 
    ncols = len(sheet_names)
    
    # Number of rows to plot
    nrows = 1
    
    alphabet = 'a'
    fig, axes = plt.subplots(nrows , ncols, figsize=(ncols*12, 8))
    
    row =0
    count = 0
    
    palette_colors = sn.color_palette()
    palette_dict = {model: color for model, color in zip(DAs, palette_colors)}
    
    if(os.path.exists(file_path)):
        for sheet_name in sheet_names:
            ax = axes[row]
            # Load the data 
            df = pd.read_excel(convergence_file_path, sheet_name=sheet_name)
            df_melted = pd.melt(df, id_vars="Iteration", value_vars=DAs, var_name="DA", value_name="Cost")
            sn.lineplot(data= df_melted, y="Cost", x='Iteration', hue='DA', ax=ax, palette= palette_dict)
            
            # Title font size
            ax.set_title('({0})-{1} Convergence'.format(alphabet, algorithms[count+1]), fontsize=18)
            
            # X and Y label font sizes
            ax.set_xlabel('Iterarion', fontsize=18)
            ax.set_ylabel('Cost', fontsize=18)
            
            # Tick label font size
            ax.tick_params(axis='both', which='major', labelsize=18)
            
            # Legend font size
            ax.legend(fontsize=16)
            
            row+=1
            
            count += 1
            alphabet = chr(ord('a') + count)
            
        save_path = os.path.join(fig_path, 'Algorithms_Convergnece.png')
        eps_path = os.path.join(fig_path, 'Algorithms_Convergnece.eps')
        
        print(save_path)
        fig.savefig(save_path,  bbox_inches="tight")
        fig.savefig(eps_path,  bbox_inches="tight", format="eps", dpi=1200)
            

convergence_file_path = os.path.join(fig_path,  "All Results_Convergence.xlsx")
plot_algorithms_convergence(convergence_file_path, algorithms, DAs ,fig_path)
        
    
    