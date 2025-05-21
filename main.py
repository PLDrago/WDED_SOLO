import os
import pandas as pd
from data_utils import load_excel_data, clean_transactions
from analysis_utils import (
    basic_stats,
    get_top_revenue_countries,
    run_apriori_analysis,
    plot_all_basic_stats,
    generate_word_report
)

NUM_COUNTRIES = 12

df = load_excel_data("zakupy-online.xlsx", sheet_name="Year 2009-2010")
print(f"Wczytano dane: {df.shape}")

df = clean_transactions(df)
print(f"Po czyszczeniu: {df.shape}")

daily_summary = basic_stats(df, save_path="outputs")

plot_all_basic_stats(df, daily_summary, save_path="outputs")

top_countries = get_top_revenue_countries(df, top_n=NUM_COUNTRIES)
print(f"\nNajbogatsze kraje ({NUM_COUNTRIES}):")
for i, country in enumerate(top_countries, start=1):
    print(f"{i}. {country}")

run_apriori_analysis(df, country_list=top_countries, output_dir="outputs/Market Basket Analysis")

generate_word_report(df, daily_summary, outputs_dir=".")
