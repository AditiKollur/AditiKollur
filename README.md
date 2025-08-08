```python
import streamlit as st
import pandas as pd
import io
import xlsxwriter

class DataReconciliationApp:
    def __init__(self):
        self.df1 = None
        self.df2 = None
        self.selected_values = None
        self.first_selected_columns = []
        self.remaining_columns = []
        self.numeric_column = None

    def load_data(self):
        uploaded_file1 = st.file_uploader("Upload Original Data (CSV or Excel)", type=["csv", "xlsx"], key="file1")
        uploaded_file2 = st.file_uploader("Upload Transformed Data (CSV or Excel)", type=["csv", "xlsx"], key="file2")

        if uploaded_file1 and uploaded_file2:
            if uploaded_file1.name.endswith('.csv'):
                self.df1 = pd.read_csv(uploaded_file1)
            else:
                self.df1 = pd.read_excel(uploaded_file1)

            if uploaded_file2.name.endswith('.csv'):
                self.df2 = pd.read_csv(uploaded_file2)
            else:
                self.df2 = pd.read_excel(uploaded_file2)

            st.success("Files loaded successfully!")

    def select_first_columns(self):
        if self.df1 is not None and self.df2 is not None:
            string_columns = [col for col in self.df1.columns if self.df1[col].dtype == 'object']
            with st.container():
                st.markdown(
                    "<div style='overflow-x:auto; overflow-y:auto; max-height:400px;'>",
                    unsafe_allow_html=True
                )
                self.first_selected_columns = st.multiselect(
                    "Select first set of string columns for grouping",
                    string_columns
                )
                st.markdown("</div>", unsafe_allow_html=True)

            if st.button("Submit First Selection"):
                self.generate_first_table()

    def generate_first_table(self):
        df1 = self.df1.copy()
        df2 = self.df2.copy()

        if not self.first_selected_columns:
            st.warning("Please select at least one column.")
            return

        df1["_filter_key"] = df1[self.first_selected_columns].astype(str).agg('_'.join, axis=1)
        df2["_filter_key"] = df2[self.first_selected_columns].astype(str).agg('_'.join, axis=1)

        numeric_cols = [col for col in df1.columns if pd.api.types.is_numeric_dtype(df1[col])]
        if numeric_cols:
            self.numeric_column = numeric_cols[0]

        grouped1 = df1.groupby("_filter_key")[self.numeric_column].sum().reset_index()
        grouped2 = df2.groupby("_filter_key")[self.numeric_column].sum().reset_index()

        merged = grouped1.merge(grouped2, on="_filter_key", suffixes=("_original", "_transformed"), how="outer").fillna(0)
        merged["status"] = merged.apply(lambda row: "OK" if row[f"{self.numeric_column}_original"] == row[f"{self.numeric_column}_transformed"]
                                        else ("Missing" if row[f"{self.numeric_column}_transformed"] == 0 else "Anomaly"), axis=1)

        st.dataframe(merged, use_container_width=True)

        self.remaining_columns = [col for col in df1.columns if col not in self.first_selected_columns and df1[col].dtype == 'object']
        self.selected_values = st.multiselect("Filter by key", merged["_filter_key"].unique())

        if st.button("Next Drill Down"):
            self.generate_second_page()

        if st.button("Export to Excel"):
            self.export_to_excel(merged)

    def generate_second_page(self):
        df1 = self.df1.copy()
        df2 = self.df2.copy()

        # NEW: Safe filtering
        if "_filter_key" in df1.columns and self.selected_values:
            df1 = df1[df1["_filter_key"].isin(self.selected_values)]
            df2 = df2[df2["_filter_key"].isin(self.selected_values)]

        with st.container():
            st.markdown(
                "<div style='overflow-x:auto; overflow-y:auto; max-height:400px;'>",
                unsafe_allow_html=True
            )
            second_selection = st.multiselect(
                "Select remaining string columns for second grouping",
                self.remaining_columns
            )
            st.markdown("</div>", unsafe_allow_html=True)

        if st.button("Submit Second Selection"):
            if not second_selection:
                st.warning("Please select at least one column.")
                return

            combined_selection = self.first_selected_columns + second_selection
            df1["_filter_key2"] = df1[combined_selection].astype(str).agg('_'.join, axis=1)
            df2["_filter_key2"] = df2[combined_selection].astype(str).agg('_'.join, axis=1)

            grouped1 = df1.groupby("_filter_key2")[self.numeric_column].sum().reset_index()
            grouped2 = df2.groupby("_filter_key2")[self.numeric_column].sum().reset_index()

            merged = grouped1.merge(grouped2, on="_filter_key2", suffixes=("_original", "_transformed"), how="outer").fillna(0)
            merged["status"] = merged.apply(lambda row: "OK" if row[f"{self.numeric_column}_original"] == row[f"{self.numeric_column}_transformed"]
                                            else ("Missing" if row[f"{self.numeric_column}_transformed"] == 0 else "Anomaly"), axis=1)

            st.dataframe(merged, use_container_width=True)

            if st.button("Export Drill Down to Excel"):
                self.export_to_excel(merged)

    def export_to_excel(self, merged_df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, sheet_name="Reconciliation", index=False)
            workbook = writer.book

            string_cols = [col for col in merged_df.columns if merged_df[col].dtype == 'object']
            num_cols = [col for col in merged_df.columns if pd.api.types.is_numeric_dtype(merged_df[col])]

            for col in string_cols:
                if col.startswith("_filter_key"):
                    chart_sheet = workbook.add_worksheet(f"Chart_{col}")
                    chart = workbook.add_chart({'type': 'column'})

                    for idx, num_col in enumerate(num_cols):
                        chart.add_series({
                            'name':       num_col,
                            'categories': ['Reconciliation', 1, merged_df.columns.get_loc(col), len(merged_df), merged_df.columns.get_loc(col)],
                            'values':     ['Reconciliation', 1, merged_df.columns.get_loc(num_col), len(merged_df), merged_df.columns.get_loc(num_col)],
                        })

                    chart.set_title({'name': f"Comparison for {col}"})
                    chart.set_x_axis({'name': col})
                    chart.set_y_axis({'name': 'Values'})
                    chart_sheet.insert_chart('B2', chart)

        st.download_button(
            label="Download Excel File",
            data=output.getvalue(),
            file_name="reconciliation_with_charts.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

def main():
    st.title("Data Reconciliation Tool with Drill Down & Charts")
    app = DataReconciliationApp()
    app.load_data()
    app.select_first_columns()

if __name__ == "__main__":
    main()
```
