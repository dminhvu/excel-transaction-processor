import streamlit as st
import pandas as pd
import io

# Define the mapping for transaction types
TRANSACTION_TYPES_MAP = {
    "INVOICE": "INV",
    "CREDIT MEMO": "CRD",
}


def process_transactions(df):
    """Process the DataFrame and extract transaction records."""
    records = []
    current_customer = None

    for idx, row in df.iterrows():
        # Convert row to string and check if it contains data
        row_data = row.astype(str)

        # Check for customer name row
        if "Customer Name" in str(row_data.iloc[2]):
            # Get the next row which should contain the customer name
            if idx + 1 < len(df):
                try:
                    next_row = df.iloc[idx + 1]
                    if not pd.isna(next_row.iloc[2]):  # Check if customer name exists
                        current_customer = str(next_row.iloc[2]).strip()
                except IndexError:
                    current_customer = None
            continue

        # Check if this is a transaction row (starts with INVOICE or CREDIT MEMO)
        doc_type = str(row_data.iloc[0]).strip().upper()
        if doc_type in TRANSACTION_TYPES_MAP:
            # Extract transaction type
            trans_type = TRANSACTION_TYPES_MAP.get(doc_type, doc_type)

            # Extract document number
            doc_number = str(row_data.iloc[1]).strip()

            # Extract and format date
            doc_date = str(row_data.iloc[4]).strip()
            # Format as DD/MM/YYYY (without time)
            try:
                doc_date = pd.to_datetime(doc_date, errors="coerce").strftime(
                    "%d/%m/%Y"
                )
                if pd.isna(doc_date):
                    doc_date = ""
            except Exception:
                doc_date = ""

            # Extract balance
            try:
                # Get the 12th column (Index 11, since it's zero-based)
                balance = str(row_data.iloc[11]).strip()
                # Remove any currency symbols, commas, and spaces
                balance = balance.replace(",", "").replace("$", "").strip()
                balance = float(balance)
                # Format as 2 decimal places
                balance = f"{balance:.2f}"
            except (ValueError, IndexError):
                balance = "0.00"

            # Create record
            if current_customer:
                records.append(
                    {
                        "Debtor Reference": current_customer,
                        "Transaction Type": trans_type,
                        "Document Number": doc_number,
                        "Document Date": doc_date,
                        "Document Balance": balance,
                    }
                )

    return records


def main():
    st.title("Excel Transaction Processor")
    st.write("""
        Upload your Excel (`.xlsx`) file to extract transaction records and download them as a CSV file.
    """)

    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        try:
            # Read the Excel file
            df = pd.read_excel(uploaded_file, dtype=str)
            st.success("Excel file successfully uploaded and read.")

            # Process the transactions
            records = process_transactions(df)
            st.write(f"**Processed {len(records)} transactions**")

            if records:
                processed_df = pd.DataFrame(records)
                st.dataframe(processed_df)

                # Convert DataFrame to CSV
                csv_buffer = io.StringIO()
                processed_df.to_csv(csv_buffer, index=False)
                csv_data = csv_buffer.getvalue()

                # Provide download button
                st.download_button(
                    label="Download Processed Data as CSV",
                    data=csv_data,
                    file_name="processed_transactions.csv",
                    mime="text/csv",
                )
            else:
                st.warning("No transaction records found in the uploaded file.")

        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")


if __name__ == "__main__":
    main()
