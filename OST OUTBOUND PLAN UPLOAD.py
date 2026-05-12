import streamlit as st
import pandas as pd
import snowflake.connector
from datetime import datetime

# --- Snowflake connection ---
@st.cache_resource
def get_snowflake_connection():
    conn = snowflake.connector.connect(
        user="lrubal@chewy.com",
        account="chewy.us-east-1",
        database="EDLDB",
        schema="SC_ORDER_ROUTING_TEAM_SANDBOX",  #SNOP_SANDBOX
        authenticator="externalbrowser",
        autocommit=False,
        role="SC_ORDER_ROUTING_TEAM_DEVELOPER",  #SNOP_DEVELOPER
        warehouse="SC_FORECAST_WH"
    )
    return conn

# --- UI ---
st.set_page_config(page_title="SNOP Forecast Upload", layout="wide")
st.title("📤 SNOP Forecast Upload")
st.markdown ("""
Upload S&OP OB 6-6 plans.

    File Requirements: 
    - The file should be .xlsx
    - Ensure the file has the sheet you want to upload. Additional sheets do not matter if they are in the file. The program will ignore them. 
        - For the OB 6-6 plan sheet name should be '6to6'
        - Upload_datetime is not required in your excel file. The program will stamp the upload time during upload automatically. 
        - Ensure column names in file match. Tables being loaded.
            - EDLDB.SC_ORDER_ROUTING_TEAM_SANDBOX.OST_SOP_OUTBOUND_PLAN_UPLOAD 

If you run into any issues, use the snowflake web UI and load the table individually.
""")
st.markdown("---")


# --- Show Latest Upload Timestamp ---
try:
    # Get connection
    conn = get_snowflake_connection()
    with conn.cursor() as cursor:
        cursor.execute("""
            SELECT 'S&OP Daily Plan OB (6-6)' as Plan, MAX(UPLOAD_DATETIME) AS LAST_UPLOAD
            FROM EDLDB.SC_ORDER_ROUTING_TEAM_SANDBOX.OST_SOP_OUTBOUND_PLAN_UPLOAD 
        """)
        results = cursor.fetchall()
        if results:
            latest_df = pd.DataFrame(results, columns=["Plan", "Last Upload"])
            st.subheader("🕒 Latest Uploads")
            st.table(latest_df)
        else:
            st.info("No uploads found yet.")
except Exception as e:
    st.error(f"❌ Could not fetch last upload times: {e}")

# --- File Upload ---
uploaded_file = st.file_uploader("📁 Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Get connection
        conn = get_snowflake_connection()
        with conn.cursor() as cursor:
            # Load Excel
            xl = pd.ExcelFile(uploaded_file)
            outbound_df = xl.parse("6to6") #OB 6-6 tab
    
            # Clean column headers
            outbound_df.columns = outbound_df.columns.str.strip().str.upper()
            
            # Convert relevant columns to just dates
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for col in ["FORECAST_DATE", "FORECAST_WEEK_BEG", "PUBLISH_WEEK"]:
                if col in outbound_df.columns:
                    outbound_df[col] = pd.to_datetime(outbound_df[col]).dt.date
    
            # Add upload timestamp
            outbound_df["UPLOAD_DATETIME"] = now
    
            # Reorder columns
            outbound_columns = [
                "UPLOAD_DATETIME",
                "FORECAST_DATE",
                "FORECAST_WEEK_BEG",
                "PUBLISH_WEEK",
                "FULFILLMENT_CENTER_NAME",
                "ALLOCATED_UNITS_SIX_TO_SIX_UNBUFFERED",
                "BUFFER_AMOUNT",
                "SOP_PLANNED_UNITS"
            ]
    
            # Validate expected columns are present
            missing_outbound = set(outbound_columns[1:]) - set(outbound_df.columns)
    
            if missing_outbound:
                st.error(f"❌ Missing columns in '6to6' tab: {missing_outbound}")
            else:
                outbound_df = outbound_df[outbound_columns]

            if not missing_outbound:
                st.subheader("✅ Preview: Outbound Forecast")
                st.dataframe(outbound_df.head(3))

                if st.button("Upload"):
                    try:
                        with st.spinner("Uploading your forecast..."):
                            #--- Upload Outbound Forecast ---
                            # 🧹 Delete rows from current week for OB 6-6
                            cursor.execute("""
                                 DELETE FROM EDLDB.SC_ORDER_ROUTING_TEAM_SANDBOX.OST_SOP_OUTBOUND_PLAN_UPLOAD
                                WHERE UPLOAD_DATETIME >= DATE_TRUNC('WEEK', CURRENT_TIMESTAMP)
                            """)
                            # Prepare data for insert
                            outbound_data = [tuple(row) for row in outbound_df.itertuples(index=False)]
            
                            #Insert Outbound Forecast
                            cursor.executemany("""
                                INSERT INTO EDLDB.SC_ORDER_ROUTING_TEAM_SANDBOX.OST_SOP_OUTBOUND_PLAN_UPLOAD  (
                                    UPLOAD_DATETIME,
                                    FORECAST_DATE,
                                    FORECAST_WEEK_BEG,
                                    PUBLISH_WEEK,
                                    FULFILLMENT_CENTER_NAME,
                                    ALLOCATED_UNITS_SIX_TO_SIX_UNBUFFERED,
                                    BUFFER_AMOUNT,
                                    SOP_PLANNED_UNITS
                                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                            """,outbound_data)
    
                            # Commit when all succeeds
                            conn.commit()                  
                        st.success("🚀 Data successfully uploaded to Snowflake!")
                        st.markdown("---")
    
                        # Display uploaded 6-6 data summary
                        ob_query = """
                                SELECT
                                UPLOAD_DATETIME,
                                FULFILLMENT_CENTER_NAME,
                                publish_week,
                                forecast_week_beg,
                                SUM(ALLOCATED_UNITS_SIX_TO_SIX_UNBUFFERED) As "Workable_Plan",
                                SUM(SOP_PLANNED_UNITS) AS "SNOP_Plan"
                                FROM EDLDB.SC_ORDER_ROUTING_TEAM_SANDBOX.OST_SOP_OUTBOUND_PLAN_UPLOAD
                                WHERE
                                UPLOAD_DATETIME = (
                                                    SELECT MAX(UPLOAD_DATETIME)
                                                        FROM EDLDB.SC_ORDER_ROUTING_TEAM_SANDBOX.OST_SOP_OUTBOUND_PLAN_UPLOAD
                                                    )
                                AND forecast_week_beg BETWEEN date_trunc('Week',current_date)-1 AND date_trunc('Week',current_date)+27
                                GROUP BY all
                                ORDER BY forecast_week_beg,FULFILLMENT_CENTER_NAME
                            """
                        cursor.execute(ob_query)
                        rows = cursor.fetchall()
                        cols = [desc[0] for desc in cursor.description]
                        
                        if rows:
                            df_pivot = pd.DataFrame(rows, columns=cols)
                            st.subheader("📊 OB 6 to 6 Upload Summary")
                            # Combined pivot with MultiIndex columns
                            # Build combined pivot with MultiIndex columns
                            pivot_combined = df_pivot.pivot_table(
                            index="FULFILLMENT_CENTER_NAME",
                            columns="FORECAST_WEEK_BEG",
                            values=["Workable_Plan", "SNOP_Plan"],
                            aggfunc="sum",
                            fill_value=0
                             )
    
                            # Optional: sort forecast weeks left-to-right
                            pivot_combined = pivot_combined.sort_index(axis=1, level=1)
                            # Flatten with "week_metric" pattern
                            pivot_combined.columns = [f"{week}_{metric}" for metric, week in pivot_combined.columns]
                        
                            # Display it
                            st.dataframe(pivot_combined)
                        else:
                            st.info("No 6 to 6 forecast found in your upload for lock weeks.")
                            
    
                    except Exception as e:
                        conn.rollback()  # 🚨 Roll back if anything fails
                        st.error(f"❌ Error: {e}")
                               
    except Exception as e:
        st.error(f"❌ Error: {e}")