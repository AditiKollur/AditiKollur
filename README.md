```\
country_list = sorted(
    set(df_cy["Managed country"].dropna().unique())
    | set(df_py["Managed country"].dropna().unique())
)

write_commentary_to_word_colored(
    commentary_dict=commentary,
    output_file="TRI_Commentary_Final.docx",
    country_names=country_list
)

