# =============================================================================
#  Algo19 Utils - Full FIFO Position Processor in Tab 0
#  All original logic unchanged - only UI moved into tab0
# =============================================================================

import streamlit as st
import pandas as pd
import numpy as np
from collections import deque
import os
import shutil
import traceback
from datetime import date
import tempfile
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Algo19 Utils", layout="wide")

# ───────────────────────────────────────────────────────────────────────────────
#   ALL ORIGINAL FUNCTIONS ── COMPLETELY UNCHANGED
# ───────────────────────────────────────────────────────────────────────────────

def safe_read_csv(path):
    """Read CSV only if file exists and is not empty"""
    if not os.path.exists(path):
        return None
    try:
        if os.path.getsize(path) == 0:
            return None  # empty file
        df = pd.read_csv(path)
        if df.empty:
            return None
        return df
    except Exception:
        return None

def load_summary(summary_path):
    df = pd.read_excel(summary_path)
    noren_user = df[df["Broker"] == "MasterTrust_Noren"]["UserID"]
    iifl_user = df[df["Broker"] == "IIFL_CDC"]["UserID"]
    other_user = df[
        (df["Broker"] != "IIFL_CDC") &
        (df["Broker"] != "MasterTrust_Noren")
    ]["UserID"]
    return df, noren_user, iifl_user, other_user


def load_orderbook(orderbook_path):
    df = pd.read_csv(orderbook_path)
    lst = df.columns.tolist()
    lst = lst[1:]
    lst.append("Status_code")
    df.columns = lst
    df = df[df["Status"] == "COMPLETE"]
    df = df[df["Product"] == "NRML"]
    return df


def filter_by_users(df, users):
    return df[df["User ID"].isin(users)]


def clean_symbols(df):
    df = df.copy()
    mask = df["Exchange"].isin(["NFO", "BFO"])
    df.loc[mask, "Symbol"] = (
        df.loc[mask, "Symbol"]
        .astype(str)
        .str.upper()
        .str.replace(" ", "", regex=False)
        .str.extract(r'(\d{5}(PE|CE)|((PE|CE)\d{5}))', expand=False)
        .iloc[:, 0]
        .str.replace(r'(PE|CE)(\d{5})', r'\2\1', regex=True)
    )
    return df


def separate_exchanges(df):
    df_nfo = df[df["Exchange"] == "NFO"].copy()
    df_bfo = df[df["Exchange"] == "BFO"].copy()
    for d in [df_nfo, df_bfo]:
        if "PNL" not in d.columns:
            d["PNL"] = 0.0
        else:
            d["PNL"] = d["PNL"].astype(float)
        if "Exit_time" not in d.columns:
            d["Exit_time"] = pd.NaT
        else:
            d["Exit_time"] = pd.to_datetime(d["Exit_time"], errors="coerce")
        if "Net_Quantity" not in d.columns:
            d["Net_Quantity"] = 0
    return df_nfo, df_bfo


def handle_carryforward(df, carry_path, users, exchange, enable_carry, broker_type):
    if not enable_carry:
        return df

    if not os.path.exists(carry_path):
        st.warning(f"[{exchange}] Carry file not found → skipping: {carry_path}")
        return df

    try:
        carry_df = pd.read_csv(carry_path)
        if carry_df.empty:
            st.info(f"[{exchange}] Carry file empty → no carry applied")
            return df

        carry_df = carry_df[carry_df["User ID"].isin(users)]
        if carry_df.empty:
            st.info(f"[{exchange}] No carry positions for these users")
            return df

        carry_df["Order Time"] = pd.to_datetime(carry_df["Order Time"], format="%d-%b-%Y %H:%M:%S", errors="coerce")
        df["Order Time"] = pd.to_datetime(df["Order Time"], format="%d-%b-%Y %H:%M:%S", errors="coerce")

        if broker_type == "Noren":
            df = pd.concat([carry_df, df], ignore_index=True) # first order of carry_df then today orders
            df = df.sort_values(by="Order Time", kind="mergesort").reset_index(drop=True)
        elif broker_type == "IIFL":
            df = pd.concat([df, carry_df], ignore_index=True) # first order of today then carry_df orders

        if "Unnamed: 0" in df.columns:
            df.drop(columns=["Unnamed: 0"], inplace=True)

        st.success(f"[{exchange}] Added {len(carry_df)} carry rows")
        return df

    except Exception as e:
        st.error(f"[{exchange}] Carry file error: {str(e)}")
        return df


def apply_fifo(df, exchange, output_final_path, output_carry_path):
    df = df.copy()

    if df.empty:
        pd.DataFrame().to_csv(output_final_path, index=False)
        pd.DataFrame().to_csv(output_carry_path, index=False)
        return pd.DataFrame(), pd.DataFrame()

    unique_users = df["User ID"].unique()
    total_realized_pnl = 0.0
    new_df = pd.DataFrame()

    for user in unique_users:
        user_df = df[df["User ID"] == user].copy()

        sell_mask = user_df["Transaction"].eq("SELL")
        user_df.loc[sell_mask, "Quantity"] = -user_df.loc[sell_mask, "Quantity"].abs()

        user_df["Exchange Time"] = pd.to_datetime(
            user_df["Exchange Time"], format="%d-%b-%Y %H:%M:%S", errors="coerce"
        )

        symbols = user_df["Symbol"].unique().tolist()
        user_new_df = pd.DataFrame()

        for symbol in symbols:
            test_df = user_df[user_df["Symbol"] == symbol].copy().reset_index(drop=True)
            if test_df.empty:
                continue

            n_rows = len(test_df)
            qty = test_df["Quantity"].astype(int).to_numpy()
            price = test_df["Avg Price"].astype(float).to_numpy()
            txn = test_df["Transaction"].to_numpy()
            t = test_df["Exchange Time"].to_numpy()

            pnl = np.zeros(n_rows, dtype=float)
            net_qty = np.zeros(n_rows, dtype=int)
            exit_time = pd.Series([pd.NaT] * n_rows, dtype="datetime64[ns]").to_numpy()
            remain = np.abs(qty).astype(int)

            if txn[0] == "SELL":
                sell_q = deque()
                for i in range(n_rows):
                    if txn[i] == "SELL":
                        sell_q.append([i, remain[i], price[i]])
                    else:
                        need = remain[i]
                        while need > 0 and sell_q:
                            s_idx, s_rem, s_px = sell_q[0]
                            matched = min(need, s_rem)
                            pnl[i] += (s_px - price[i]) * matched
                            need -= matched
                            s_rem -= matched
                            if s_rem == 0:
                                sell_q.popleft()
                            else:
                                sell_q[0][1] = s_rem
                        net_qty[i] = need
                        if need == 0:
                            exit_time[i] = t[i]
                for s_idx, s_rem, _ in sell_q:
                    net_qty[s_idx] = -s_rem
            else:
                buy_q = deque()
                for i in range(n_rows):
                    if txn[i] == "BUY":
                        buy_q.append([i, remain[i], price[i]])
                    else:
                        need = remain[i]
                        while need > 0 and buy_q:
                            b_idx, b_rem, b_px = buy_q[0]
                            matched = min(need, b_rem)
                            pnl[i] += (price[i] - b_px) * matched
                            need -= matched
                            b_rem -= matched
                            if b_rem == 0:
                                exit_time[b_idx] = t[i]
                                buy_q.popleft()
                            else:
                                buy_q[0][1] = b_rem
                        net_qty[i] = -need
                for b_idx, b_rem, _ in buy_q:
                    net_qty[b_idx] = b_rem

            test_df["PNL"] = pnl
            test_df["Net_Quantity"] = net_qty
            test_df["Exit_time"] = exit_time

            user_new_df = pd.concat([user_new_df, test_df], ignore_index=True)
            total_realized_pnl += float(pnl.sum())

        new_df = pd.concat([new_df, user_new_df], ignore_index=True)

    carry_fwd_pos_df = new_df[new_df["Net_Quantity"] != 0].copy()

    new_df.to_csv(output_final_path, index=False)
    carry_fwd_pos_df.to_csv(output_carry_path, index=False)

    st.write(f"[{exchange}] Total Realized PNL: **{total_realized_pnl:,.2f}**")

    return new_df, carry_fwd_pos_df


def create_position_summary(df_all):
    if df_all.empty:
        return pd.DataFrame()

    grouped = df_all.groupby(['User ID', 'Exchange', 'Symbol'])
    rows = []

    for (user_id, exchange, symbol), group in grouped:
        buy_mask = group['Quantity'] > 0
        sell_mask = group['Quantity'] < 0

        buy_qty = group.loc[buy_mask, 'Quantity'].sum()
        buy_value = (group.loc[buy_mask, 'Quantity'] * group.loc[buy_mask, 'Avg Price']).sum()
        buy_avg = buy_value / buy_qty if buy_qty != 0 else 0.0

        sell_qty = abs(group.loc[sell_mask, 'Quantity'].sum())
        sell_value = abs((group.loc[sell_mask, 'Quantity'] * group.loc[sell_mask, 'Avg Price']).sum())
        sell_avg = sell_value / sell_qty if sell_qty != 0 else 0.0

        net_qty = group['Net_Quantity'].sum()
        realized = group['PNL'].sum()

        if buy_qty == 0 and sell_qty == 0 and net_qty == 0 and realized == 0:
            continue

        rows.append({
            'UserID': user_id,
            'Product': 'NRML',
            'Exchange': exchange,
            'Symbol': symbol,
            'Net Qty': int(net_qty),
            'Buy Qty': int(buy_qty),
            'Buy Avg Price': round(buy_avg, 2),
            'Buy Value': round(buy_value, 2),
            'Sell Qty': int(sell_qty),
            'Sell Avg Price': round(sell_avg, 2),
            'Sell Value': round(sell_value, 2),
            'Carry Fwd Qty': 0,
            'Realized Profit': round(realized, 2),
            'Unrealized Profit': 0.0,
            'P&L': round(realized, 2),
        })

    if not rows:
        return pd.DataFrame()

    summary = pd.DataFrame(rows)
    summary = summary.sort_values(['UserID', 'Exchange', 'Symbol']).reset_index(drop=True)
    summary.insert(0, 'S.No.', range(1, len(summary) + 1))
    return summary


def update_eod_positions(eod_path, position_summary, output_path):
    if position_summary.empty:
        if os.path.normpath(eod_path) != os.path.normpath(output_path):
            shutil.copy(eod_path, output_path)
        return output_path

    eod = pd.read_csv(eod_path)

    eod["Symbol_clean"] = (
        eod["Symbol"].astype(str).str.upper().str.replace(" ", "", regex=False)
        .str.extract(r'(\d{5}(PE|CE)|((PE|CE)\d{5}))', expand=False).iloc[:, 0]
        .str.replace(r'(PE|CE)(\d{5})', r'\2\1', regex=True).fillna("")
    )

    eod["key"] = eod["UserID"].astype(str).str.strip() + "|" + eod["Symbol_clean"].str.strip()

    position_summary["key"] = (
        position_summary["UserID"].astype(str).str.strip() + "|" +
        position_summary["Symbol"].str.strip()
    )

    cols = ["Net Qty", "Buy Qty", "Buy Avg Price", "Buy Value",
            "Sell Qty", "Sell Avg Price", "Sell Value",
            "Carry Fwd Qty", "Realized Profit", "Unrealized Profit", "P&L"]
    cols = [c for c in cols if c in eod.columns]

    merged = eod.merge(position_summary[["key"] + cols], on="key", how="left", suffixes=("", "_new"))

    for col in cols:
        merged[col] = merged[f"{col}_new"].combine_first(merged[col])

    drop = [f"{c}_new" for c in cols] + ["key", "Symbol_clean"]
    drop = [c for c in drop if c in merged.columns]
    updated = merged.drop(columns=drop)

    updated.to_csv(output_path, index=False)
    return output_path


def process_broker(orderbook_df, users, broker_type, enable_carry,
                   carry_nfo_path, carry_bfo_path,
                   final_nfo_path, carry_out_nfo_path,
                   final_bfo_path, carry_out_bfo_path):
    if users.empty:
        return pd.DataFrame()

    df = filter_by_users(orderbook_df, users)
    df = clean_symbols(df)
    df_nfo, df_bfo = separate_exchanges(df)

    df_nfo = handle_carryforward(df_nfo, carry_nfo_path, users, "NFO", enable_carry, broker_type)
    df_bfo = handle_carryforward(df_bfo, carry_bfo_path, users, "BFO", enable_carry, broker_type)

    new_nfo, carry_nfo = apply_fifo(df_nfo, "NFO", final_nfo_path, carry_out_nfo_path)
    new_bfo, carry_bfo = apply_fifo(df_bfo, "BFO", final_bfo_path, carry_out_bfo_path)

    df_all = pd.concat([new_nfo, new_bfo], ignore_index=True)
    position_summary = create_position_summary(df_all)
    return position_summary


def save_uploaded_temp(uploaded_file):
    if uploaded_file is None:
        return None
    suffix = os.path.splitext(uploaded_file.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getvalue())
        return tmp.name


def move_to_downloads(src_path, desired_name):
    if not src_path or not os.path.exists(src_path):
        return None
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(downloads, exist_ok=True)
    dest = os.path.join(downloads, desired_name)
    shutil.move(src_path, dest)
    return dest


def run_processing(
    summary_path, orderbook_path, eod_path, output_name,
    carry_nfo_iifl, carry_bfo_iifl, carry_nfo_noren, carry_bfo_noren,
    enable_iifl, enable_noren
):
    today = date.today().strftime("%Y%m%d")

    # Temp filenames used internally by apply_fifo (unchanged)
    fn_nfo_i = f"final_nfo_iifl.csv"
    cn_nfo_i = f"carry_nfo_iifl.csv"
    fb_nfo_i = f"final_bfo_iifl.csv"
    cb_nfo_i = f"carry_bfo_iifl.csv"

    fn_nfo_n = f"final_nfo_noren.csv"
    cn_nfo_n = f"carry_nfo_noren.csv"
    fb_nfo_n = f"final_bfo_noren.csv"
    cb_nfo_n = f"carry_bfo_noren.csv"

    with st.status("Processing...", expanded=True) as status:
        status.write("Loading files...")
        _, noren_users, iifl_users, _ = load_summary(summary_path)
        orderbook_df = load_orderbook(orderbook_path)

        # Load EOD once
        try:
            eod_df = pd.read_csv(eod_path)
        except Exception as e:
            st.error(f"Cannot read EOD file: {str(e)}")
            return None, None, None, None, None, None, None, None, None

        if orderbook_df.empty:
            status.update(label="No valid orders found", state="complete")
            # Return original EOD as updated (no changes)
            return eod_df, None, None, None, None, None, None, None, None

        current_df = eod_df.copy()

        # We'll collect DataFrames for download
        download_data = {
            "updated_df": None,
            "i_final_nfo": None, "i_carry_nfo": None,
            "i_final_bfo": None, "i_carry_bfo": None,
            "n_final_nfo": None, "n_carry_nfo": None,
            "n_final_bfo": None, "n_carry_bfo": None,
        }

        if not iifl_users.empty:
            status.write("Processing IIFL...")
            pos_i = process_broker(
                orderbook_df, iifl_users, "IIFL", enable_iifl,
                carry_nfo_iifl, carry_bfo_iifl,
                fn_nfo_i, cn_nfo_i, fb_nfo_i, cb_nfo_i
            )
            if not pos_i.empty:
                current_df = update_eod_positions_df(current_df, pos_i)  # ← using the df version
                download_data["updated_df"] = current_df.copy()

            # Read what apply_fifo wrote
            # Then use it:
            download_data["i_final_nfo"]  = safe_read_csv(fn_nfo_i)
            download_data["i_carry_nfo"]  = safe_read_csv(cn_nfo_i)
            download_data["i_final_bfo"]  = safe_read_csv(fb_nfo_i)
            download_data["i_carry_bfo"]  = safe_read_csv(cb_nfo_i)

            download_data["n_final_nfo"]  = safe_read_csv(fn_nfo_n)
            download_data["n_carry_nfo"]  = safe_read_csv(cn_nfo_n)
            download_data["n_final_bfo"]  = safe_read_csv(fb_nfo_n)
            download_data["n_carry_bfo"]  = safe_read_csv(cb_nfo_n)

        if not noren_users.empty:
            status.write("Processing Noren...")
            pos_n = process_broker(
                orderbook_df, noren_users, "Noren", enable_noren,
                carry_nfo_noren, carry_bfo_noren,
                fn_nfo_n, cn_nfo_n, fb_nfo_n, cb_nfo_n
            )
            if not pos_n.empty:
                current_df = update_eod_positions_df(current_df, pos_n)
                download_data["updated_df"] = current_df.copy()

            if os.path.exists(fn_nfo_n):
                download_data["n_final_nfo"] = pd.read_csv(fn_nfo_n)
            if os.path.exists(cn_nfo_n):
                download_data["n_carry_nfo"] = pd.read_csv(cn_nfo_n)
            if os.path.exists(fb_nfo_n):
                download_data["n_final_bfo"] = pd.read_csv(fb_nfo_n)
            if os.path.exists(cb_nfo_n):
                download_data["n_carry_bfo"] = pd.read_csv(cb_nfo_n)

        status.update(label="Done", state="complete")

    return download_data

def update_eod_positions_df(eod_df, position_summary):
    if position_summary.empty:
        return eod_df.copy()

    eod = eod_df.copy()

    eod["Symbol_clean"] = (
        eod["Symbol"].astype(str).str.upper().str.replace(" ", "", regex=False)
        .str.extract(r'(\d{5}(PE|CE)|((PE|CE)\d{5}))', expand=False).iloc[:, 0]
        .str.replace(r'(PE|CE)(\d{5})', r'\2\1', regex=True).fillna("")
    )

    eod["key"] = eod["UserID"].astype(str).str.strip() + "|" + eod["Symbol_clean"].str.strip()

    position_summary["key"] = (
        position_summary["UserID"].astype(str).str.strip() + "|" +
        position_summary["Symbol"].str.strip()
    )

    cols = ["Net Qty", "Buy Qty", "Buy Avg Price", "Buy Value",
            "Sell Qty", "Sell Avg Price", "Sell Value",
            "Carry Fwd Qty", "Realized Profit", "Unrealized Profit", "P&L"]
    cols = [c for c in cols if c in eod.columns]

    merged = eod.merge(position_summary[["key"] + cols], on="key", how="left", suffixes=("", "_new"))

    for col in cols:
        merged[col] = merged[f"{col}_new"].combine_first(merged[col])

    drop = [f"{c}_new" for c in cols] + ["key", "Symbol_clean"]
    drop = [c for c in drop if c in merged.columns]
    updated = merged.drop(columns=drop)

    return updated

# ───────────────────────────────────────────────────────────────────────────────
#   MAIN APP - TABS
# ───────────────────────────────────────────────────────────────────────────────

st.title("Algo19 Utils")

tab0, tab1, tab2 = st.tabs([
    "Tab 0: Noren & IIFL FIFO PNL + Positions",
    "Tab 1: Noren Bhavcopy Carryforward Update",
    "Tab 2: Jainam Calculation for Algo19"
])

# ───────────────────────────────────────────────────────────────────────────────
#   TAB 0 - FULL FIFO PROCESSOR UI
# ───────────────────────────────────────────────────────────────────────────────

with tab0:
    st.header("Noren & IIFL FIFO PNL + Position Processor (NRML / NFO / BFO)")
    st.caption("Supports separate carry-forward for IIFL & Noren • Updates EOD • Direct browser downloads (no disk path)")

    with st.form("process_form_tab0"):

        st.subheader("Input Files")

        colA, colB = st.columns(2)
        with colA:
            summary_upl   = st.file_uploader("Summary (Excel) – Broker & UserID", type=["xlsx","xls"], key="sum_tab0")
            orderbook_upl = st.file_uploader("Orderbook (CSV)", type="csv", key="ord_tab0")
        with colB:
            eod_upl       = st.file_uploader("Current EOD Positions (CSV)", type="csv", key="eod_tab0")
            output_name   = st.text_input("Suggested filename for updated positions", "updated_positions.csv", key="outname_tab0")

        st.subheader("Carry-forward Files (optional – from previous day)")

        today_str = date.today().strftime("%Y-%m-%d")

        c1, c2 = st.columns(2)
        with c1:
            carry_nfo_iifl_upl  = st.file_uploader(f"Carry NFO IIFL ({today_str})", type="csv", key="cni_tab0")
            carry_bfo_iifl_upl  = st.file_uploader(f"Carry BFO IIFL ({today_str})", type="csv", key="cbi_tab0")
        with c2:
            carry_nfo_noren_upl = st.file_uploader(f"Carry NFO Noren ({today_str})", type="csv", key="cnn_tab0")
            carry_bfo_noren_upl = st.file_uploader(f"Carry BFO Noren ({today_str})", type="csv", key="cbn_tab0")

        st.subheader("Carry-forward Options")
        enable_iifl  = st.checkbox("Enable carry-forward for IIFL", value=False, key="en_iifl_tab0")
        enable_noren = st.checkbox("Enable carry-forward for Noren", value=False, key="en_noren_tab0")

        start_btn = st.form_submit_button("START PROCESSING", type="primary", use_container_width=True)


    if start_btn:

        if not all([summary_upl, orderbook_upl, eod_upl]):
            st.error("Please upload Summary, Orderbook and EOD Positions files.")
        else:
            temp_paths = {
                "summary":      save_uploaded_temp(summary_upl),
                "orderbook":    save_uploaded_temp(orderbook_upl),
                "eod":          save_uploaded_temp(eod_upl),
                "carry_nfo_iifl":  save_uploaded_temp(carry_nfo_iifl_upl) if carry_nfo_iifl_upl else None,
                "carry_bfo_iifl":  save_uploaded_temp(carry_bfo_iifl_upl) if carry_bfo_iifl_upl else None,
                "carry_nfo_noren": save_uploaded_temp(carry_nfo_noren_upl) if carry_nfo_noren_upl else None,
                "carry_bfo_noren": save_uploaded_temp(carry_bfo_noren_upl) if carry_bfo_noren_upl else None,
            }

            with st.spinner("Processing trades (may take several seconds)..."):
                try:
                    download_data = run_processing(
                        temp_paths["summary"],
                        temp_paths["orderbook"],
                        temp_paths["eod"],
                        output_name,
                        temp_paths["carry_nfo_iifl"] or "",
                        temp_paths["carry_bfo_iifl"] or "",
                        temp_paths["carry_nfo_noren"] or "",
                        temp_paths["carry_bfo_noren"] or "",
                        enable_iifl,
                        enable_noren
                    )

                    st.success("Processing completed!")

                    st.subheader("Download Results")

                    cols = st.columns(3)

                    # Helper function to create download button safely
                    def safe_download(df, filename, label, col):
                        if df is not None and not df.empty:
                            buf = BytesIO()
                            df.to_csv(buf, index=False)
                            buf.seek(0)
                            col.download_button(
                                label=label,
                                data=buf,
                                file_name=filename,
                                mime="text/csv",
                                key=f"dl_{filename.replace('.csv','')}"
                            )

                    today = date.today().strftime("%Y%m%d")

                    with cols[0]:
                        safe_download(download_data["updated_df"], output_name, "📥 Updated Positions (final)", cols[0])
                        safe_download(download_data["i_final_nfo"], f"final_nfo_iifl.csv", "IIFL NFO Final", cols[0])
                        safe_download(download_data["n_final_nfo"], f"final_nfo_noren.csv", "Noren NFO Final", cols[0])

                    with cols[1]:
                        safe_download(download_data["i_carry_nfo"], f"carry_nfo_iifl.csv", "IIFL NFO Carry", cols[1])
                        safe_download(download_data["n_carry_nfo"], f"carry_nfo_noren.csv", "Noren NFO Carry", cols[1])
                        safe_download(download_data["i_final_bfo"], f"final_bfo_iifl.csv", "IIFL BFO Final", cols[1])

                    with cols[2]:
                        safe_download(download_data["n_final_bfo"], f"final_bfo_noren.csv", "Noren BFO Final", cols[2])
                        safe_download(download_data["i_carry_bfo"], f"carry_bfo_iifl.csv", "IIFL BFO Carry", cols[2])
                        safe_download(download_data["n_carry_bfo"], f"carry_bfo_noren.csv", "Noren BFO Carry", cols[2])

                except Exception as e:
                    st.error("Error during processing")
                    with st.expander("Show detailed error"):
                        st.code(traceback.format_exc(), language="python")

                finally:
                    # Clean up all temporary files
                    for path in temp_paths.values():
                        if path and os.path.exists(path):
                            try:
                                os.unlink(path)
                            except:
                                pass

                    # Clean files created by apply_fifo
                    for fname in [
                        f"final_nfo_iifl.csv", f"carry_nfo_iifl.csv",
                        f"final_bfo_iifl.csv", f"carry_bfo_iifl.csv",
                        f"final_nfo_noren.csv", f"carry_nfo_noren.csv",
                        f"final_bfo_noren.csv", f"carry_bfo_noren.csv",
                    ]:
                        if os.path.exists(fname):
                            try:
                                os.unlink(fname)
                            except:
                                pass


# ───────────────────────────────────────────────────────────────────────────────
#   TAB 1 & TAB 2  (placeholders – add your previous code here)
# ───────────────────────────────────────────────────────────────────────────────
# =============================================================================
#  Tab 1: Noren Bhavcopy Carryforward Update
#  (full logic from your provided code – unchanged)
# =============================================================================

# ───────────────────────────────────────────────────────────────────────────────
#   HELPERS (exactly as you provided)
# ───────────────────────────────────────────────────────────────────────────────

def save_uploaded_temp(uploaded_file):
    if uploaded_file is None:
        return None
    suffix = os.path.splitext(uploaded_file.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getvalue())
        return tmp.name


def move_to_downloads_with_updated_name(src_path, original_name):
    if not src_path or not os.path.exists(src_path):
        return None, None
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(downloads, exist_ok=True)
    base, ext = os.path.splitext(original_name)
    new_name = f"{base}_updated{ext}"
    dest = os.path.join(downloads, new_name)
    shutil.move(src_path, dest)
    return dest, new_name


def clean_symbol(symbol_series):
    return symbol_series.astype(str).str.upper().str.strip()


def load_and_prepare_bhavcopy(bhav_path, exchange, selected_expiry):
    if not bhav_path or not selected_expiry:
        return None

    try:
        df_bhav = pd.read_csv(bhav_path)

        price_col_candidates = ["SETTLEMENT", "SETTLE_PR", "Close Price", "CLOSE"]
        price_col = next((col for col in price_col_candidates if col in df_bhav.columns), None)
        if price_col is None:
            st.error(f"Could not find price column in {exchange} bhavcopy.")
            st.write("Available columns:", df_bhav.columns.tolist())
            return None
        df_bhav = df_bhav.rename(columns={price_col: "Bhav_Price"})

        if exchange == "NFO":
            if "CONTRACT_D" not in df_bhav.columns:
                st.error("NFO bhavcopy missing 'CONTRACT_D' column.")
                return None

            df_bhav["Date"] = df_bhav["CONTRACT_D"].str.extract(r'(\d{2}-[A-Z]{3}-\d{4})')
            df_bhav["Symbol"] = df_bhav["CONTRACT_D"].str.extract(r'^(.*?)(\d{2}-[A-Z]{3}-\d{4})')[0]
            df_bhav["Strike_Type"] = df_bhav["CONTRACT_D"].str.extract(r'(PE\d+|CE\d+)$')
            df_bhav["Date"] = pd.to_datetime(df_bhav["Date"], format="%d-%b-%Y", errors="coerce")
            df_bhav["Strike_Type"] = df_bhav["Strike_Type"].str.replace(
                r'^(PE|CE)(\d+)$', r'\2\1', regex=True
            )

            target_symbol = "OPTIDXNIFTY"
            df_bhav = df_bhav[
                (df_bhav["Date"].dt.date == selected_expiry) &
                (df_bhav["Symbol"] == target_symbol)
            ]

            if df_bhav.empty:
                st.warning(f"No data for expiry {selected_expiry} and {target_symbol} in NFO.")
                return None

            key_col = "Strike_Type"
            df_bhav[key_col] = clean_symbol(df_bhav[key_col])

        elif exchange == "BFO":
            if "Expiry Date" not in df_bhav.columns or "Series Code" not in df_bhav.columns:
                st.error("BFO bhavcopy missing required columns (Expiry Date, Series Code).")
                return None

            df_bhav["Date"] = pd.to_datetime(df_bhav.get("Market Summary Date", pd.Series()), format="%d %b %Y", errors="coerce")
            df_bhav["Expiry Date"] = pd.to_datetime(df_bhav["Expiry Date"], format="%d %b %Y", errors="coerce")
            df_bhav["Symbols"] = df_bhav["Series Code"].astype(str).str[-7:].str.strip()

            df_bhav = df_bhav[df_bhav["Expiry Date"].dt.date == selected_expiry]

            if df_bhav.empty:
                st.warning(f"No data for expiry {selected_expiry} in BFO.")
                return None

            key_col = "Symbols"
            df_bhav[key_col] = clean_symbol(df_bhav[key_col])

        else:
            return None

        df_bhav = df_bhav[[key_col, "Bhav_Price"]].dropna(subset=[key_col])
        df_bhav = df_bhav.drop_duplicates(subset=key_col, keep="last")
        price_map = df_bhav.set_index(key_col)["Bhav_Price"].to_dict()

        return price_map

    except Exception as e:
        st.error(f"Error processing {exchange} bhavcopy: {str(e)}")
        return None


# ───────────────────────────────────────────────────────────────────────────────
#   TAB 1 UI & PROCESSING
# ───────────────────────────────────────────────────────────────────────────────

with tab1:
    st.header("Noren Bhavcopy Carryforward Update")
    st.caption("Updates **Avg Price** in carry-forward files using NFO/BFO bhavcopy • NIFTY & SENSEX • Saves with _updated suffix to Downloads")
    st.info("If you upload carry for NFO, you must also upload NFO bhavcopy (same for BFO). You can process one, both, or none.")

    with st.form("update_carry_form_tab1"):

        st.subheader("Files & Expiry Dates")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("**NFO (NIFTY options/futures)**")
            carry_nfo_upl_tab1 = st.file_uploader("Carry-forward NFO CSV", type="csv", key="carry_nfo_tab1")
            bhav_nfo_upl_tab1  = st.file_uploader("Bhavcopy NFO CSV", type="csv", key="bhav_nfo_tab1")
            expiry_nfo_tab1    = st.date_input("NFO Expiry Date", value=date.today(), key="exp_nfo_tab1")

        with col2:
            st.markdown("**BFO (SENSEX options/futures)**")
            carry_bfo_upl_tab1 = st.file_uploader("Carry-forward BFO CSV", type="csv", key="carry_bfo_tab1")
            bhav_bfo_upl_tab1  = st.file_uploader("Bhavcopy BFO CSV", type="csv", key="bhav_bfo_tab1")
            expiry_bfo_tab1    = st.date_input("BFO Expiry Date", value=date.today(), key="exp_bfo_tab1")

        submit_tab1 = st.form_submit_button("UPDATE AVG PRICE → SAVE _updated", type="primary", use_container_width=True)


    if submit_tab1:

        has_nfo_carry = carry_nfo_upl_tab1 is not None
        has_bfo_carry = carry_bfo_upl_tab1 is not None
        has_nfo_bhav  = bhav_nfo_upl_tab1  is not None
        has_bfo_bhav  = bhav_bfo_upl_tab1  is not None

        if (has_nfo_carry and not has_nfo_bhav) or (has_bfo_carry and not has_bfo_bhav):
            st.error("Missing bhavcopy: If you upload carry-forward for NFO/BFO, you must also upload the corresponding bhavcopy.")
        elif not has_nfo_carry and not has_bfo_carry:
            st.warning("No carry-forward files uploaded. Nothing to process.")
        else:
            temp_paths = {}
            if has_nfo_carry: temp_paths["carry_nfo"] = save_uploaded_temp(carry_nfo_upl_tab1)
            if has_nfo_bhav:  temp_paths["bhav_nfo"]  = save_uploaded_temp(bhav_nfo_upl_tab1)
            if has_bfo_carry: temp_paths["carry_bfo"] = save_uploaded_temp(carry_bfo_upl_tab1)
            if has_bfo_bhav:  temp_paths["bhav_bfo"]  = save_uploaded_temp(bhav_bfo_upl_tab1)

            with st.status("Updating carry-forward files...", expanded=True) as status:

                price_map_nfo = load_and_prepare_bhavcopy(temp_paths.get("bhav_nfo"), "NFO", expiry_nfo_tab1) if "bhav_nfo" in temp_paths else None
                price_map_bfo = load_and_prepare_bhavcopy(temp_paths.get("bhav_bfo"), "BFO", expiry_bfo_tab1) if "bhav_bfo" in temp_paths else None

                updated_files = []

                # ── NFO ────────────────────────────────────────────────────────
                if "carry_nfo" in temp_paths:
                    status.write("Processing NFO carry...")
                    try:
                        df_nfo = pd.read_csv(temp_paths["carry_nfo"])
                        if "Symbol" not in df_nfo.columns or "Avg Price" not in df_nfo.columns:
                            st.error("NFO carry file missing 'Symbol' or 'Avg Price' column.")
                        else:
                            df_nfo["Symbol_clean"] = clean_symbol(df_nfo["Symbol"])
                            if price_map_nfo:
                                updated_count = df_nfo["Symbol_clean"].isin(price_map_nfo.keys()).sum()
                                df_nfo["Avg Price"] = df_nfo["Symbol_clean"].map(price_map_nfo).fillna(df_nfo["Avg Price"])
                                st.success(f"NFO: Updated {updated_count} positions with settlement price.")
                            else:
                                st.warning("No valid NFO bhavcopy data → Avg Price unchanged.")
                            df_nfo.drop(columns=["Symbol_clean"], errors="ignore", inplace=True)

                            # Overwrite temp file with updated content
                            df_nfo.to_csv(temp_paths["carry_nfo"], index=False)
                            out_path, out_name = move_to_downloads_with_updated_name(temp_paths["carry_nfo"], carry_nfo_upl_tab1.name)
                            if out_path:
                                updated_files.append((out_name, out_path))
                    except Exception as e:
                        st.error(f"NFO processing failed: {str(e)}")

                # ── BFO ────────────────────────────────────────────────────────
                if "carry_bfo" in temp_paths:
                    status.write("Processing BFO carry...")
                    try:
                        df_bfo = pd.read_csv(temp_paths["carry_bfo"])
                        if "Symbol" not in df_bfo.columns or "Avg Price" not in df_bfo.columns:
                            st.error("BFO carry file missing 'Symbol' or 'Avg Price' column.")
                        else:
                            df_bfo["Symbol_clean"] = clean_symbol(df_bfo["Symbol"])
                            if price_map_bfo:
                                updated_count = df_bfo["Symbol_clean"].isin(price_map_bfo.keys()).sum()
                                df_bfo["Avg Price"] = df_bfo["Symbol_clean"].map(price_map_bfo).fillna(df_bfo["Avg Price"])
                                st.success(f"BFO: Updated {updated_count} positions with settlement price.")
                            else:
                                st.warning("No valid BFO bhavcopy data → Avg Price unchanged.")
                            df_bfo.drop(columns=["Symbol_clean"], errors="ignore", inplace=True)

                            df_bfo.to_csv(temp_paths["carry_bfo"], index=False)
                            out_path, out_name = move_to_downloads_with_updated_name(temp_paths["carry_bfo"], carry_bfo_upl_tab1.name)
                            if out_path:
                                updated_files.append((out_name, out_path))
                    except Exception as e:
                        st.error(f"BFO processing failed: {str(e)}")

                status.update(label="Finished", state="complete")

            if updated_files:
                st.success(f"Updated files saved to your Downloads folder!")
                for fname, fpath in updated_files:
                    if os.path.exists(fpath):
                        with open(fpath, "rb") as f:
                            st.download_button(
                                label=f"📥 Download {fname}",
                                data=f,
                                file_name=fname,
                                mime="text/csv",
                                key=f"dl_{fname}"
                            )

            # Cleanup temporary files
            for p in temp_paths.values():
                if p and os.path.exists(p):
                    try:
                        os.unlink(p)
                    except:
                        pass

# =============================================================================
#  Tab 2: Jainam Calculation for Algo19
#  (full logic from your provided code – unchanged)
# =============================================================================

import re
from io import BytesIO

# ────────────────────────────────────────────────
#  Session state keys prefixed for Tab 2
# ────────────────────────────────────────────────
if 'tab2_processed' not in st.session_state:
    st.session_state.tab2_processed = False
if 'tab2_pnl_bytes' not in st.session_state:
    st.session_state.tab2_pnl_bytes = None
if 'tab2_hedge_bytes' not in st.session_state:
    st.session_state.tab2_hedge_bytes = None
if 'tab2_user_hedge_bytes' not in st.session_state:
    st.session_state.tab2_user_hedge_bytes = None


# ────────────────────────────────────────────────
#   TAB 2 UI & PROCESSING
# ────────────────────────────────────────────────

with tab2:
    st.header("Jainam Calculation for Algo19")
    st.caption("Processes SUMMARY.xlsx (MultiLeg Orders) + Orderbook CSVs → NF/SN PNL + Hedge Cost + User-wise summary")
    st.info("Upload VS20 SUMMARY.xlsx + one or more Orderbook CSV files → get downloadable results")

    col1, col2 = st.columns([5, 5])

    with col1:
        excel_file_tab2 = st.file_uploader(
            "Upload VS20 ... SUMMARY.xlsx",
            type=["xlsx", "xls"],
            key="excel_uploader_tab2"
        )

    with col2:
        csv_files_tab2 = st.file_uploader(
            "Upload Orderbook CSV files (multiple allowed)",
            type=["csv"],
            accept_multiple_files=True,
            key="csv_uploader_tab2"
        )

    if st.button("Process Jainam Files", type="primary", use_container_width=True, key="process_btn_tab2"):
        if not excel_file_tab2:
            st.error("Please upload the SUMMARY Excel file.")
        elif not csv_files_tab2:
            st.error("Please upload at least one Orderbook CSV file.")
        else:
            with st.spinner("Processing Jainam files... (usually 5–25 seconds depending on size)"):
                try:
                    # ───── Excel part ─────
                    df = pd.read_excel(excel_file_tab2, sheet_name="MultiLeg Orders")

                    # Leg-level PNL
                    leg_results = []
                    for (user, port), g in df.groupby(["User ID", "Portfolio Name"]):
                        for leg, leg_df in g.groupby("Leg ID"):
                            buy_p  = leg_df[leg_df["Transaction"] == "BUY"]["Avg Price"].mean() or 0
                            sell_p = leg_df[leg_df["Transaction"] == "SELL"]["Avg Price"].mean() or 0
                            buy_q  = leg_df[leg_df["Transaction"] == "BUY"]["Quantity"].sum() or 0
                            sell_q = leg_df[leg_df["Transaction"] == "SELL"]["Quantity"].sum() or 0
                            pnl    = (sell_p - buy_p) * sell_q

                            under = "NF" if "_NF_" in port else "SN" if "_SN_" in port else "OTHER"
                            dte   = re.search(r'(\d+DTE)', port).group(1) if re.search(r'(\d+DTE)', port) else "UNKNOWN"

                            leg_results.append({
                                "User ID": user, "Portfolio Name": port, "Underlying": under,
                                "DTE": dte, "Leg ID": leg, "Buy Avg Price": buy_p, "Sell Avg Price": sell_p,
                                "Buy Quantity": buy_q, "Sell Quantity": sell_q, "PNL": pnl
                            })

                    leg_df = pd.DataFrame(leg_results)
                    agg_df = leg_df.groupby(["User ID", "Underlying", "DTE"])["PNL"].sum().reset_index()

                    # 916 dates
                    df["Order Time"] = pd.to_datetime(df["Order Time"], errors='coerce')
                    results_916 = []
                    for (user, port), g in df.groupby(["User ID", "Portfolio Name"]):
                        if ("_NF_" not in port and "_SN_" not in port) or "916" not in port:
                            continue
                        underlying = "NF" if "_NF_" in port else "SN"
                        dte   = re.search(r'(\d+DTE)', port).group(1) if re.search(r'(\d+DTE)', port) else "UNKNOWN"
                        strike = re.search(r'_(\d{3,5})_', port).group(1) if re.search(r'_(\d{3,5})_', port) else "UNKNOWN"
                        order_date = g["Order Time"].min().date()
                        results_916.append({"User ID": user, "Underlying": underlying, "DTE": dte, "Strike": strike, "Order Date": order_date})

                    date_df = pd.DataFrame(results_916)[["DTE", "Order Date"]].drop_duplicates("DTE").rename(columns={"Order Date": "Date"})
                    agg_df = agg_df.merge(date_df, on="DTE", how="left")

                    # Save PNL to bytes
                    pnl_buf = BytesIO()
                    with pd.ExcelWriter(pnl_buf, engine='openpyxl') as w:
                        agg_df.to_excel(w, index=False)
                    pnl_buf.seek(0)
                    st.session_state.tab2_pnl_bytes = pnl_buf.getvalue()

                    # ───── CSV + Hedge part ─────
                    dfs = []
                    for f in csv_files_tab2:
                        dfc = pd.read_csv(f)
                        if len(dfc.columns) > 0:
                            dfc.columns = list(dfc.columns[1:]) + ["Tags"]
                        dfs.append(dfc)

                    final_df_csv = pd.concat(dfs, ignore_index=True)

                    # unmatched
                    final_df_csv.columns = final_df_csv.columns.str.strip()
                    final_df_csv['Order ID'] = final_df_csv['Order ID'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

                    df_excel = pd.read_excel(excel_file_tab2, sheet_name="MultiLeg Orders")
                    df_excel.columns = df_excel.columns.str.strip()
                    df_excel['Order ID'] = df_excel['Order ID'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

                    unmatched = final_df_csv[~final_df_csv['Order ID'].isin(df_excel['Order ID'])]

                    # Hedge calculation
                    df = unmatched[unmatched['Product'] == 'NRML'].copy()
                    df['Product'] = df['Product'].astype(str).str.strip()
                    df['Transaction'] = df['Transaction'].str.strip().str.upper()
                    df["total_value"] = df["Avg Price"] * df["Quantity"]

                    grouped = df.groupby(['User ID', 'Symbol', 'Transaction'], as_index=False).agg({
                        'total_value': 'sum',
                        'Quantity': 'sum'
                    }).rename(columns={'total_value': 'price'})

                    grouped['avg_price'] = grouped['price'] / grouped['Quantity'].replace(0, pd.NA)

                    pivot = grouped.pivot(
                        index=['User ID', 'Symbol'],
                        columns='Transaction',
                        values=['avg_price', 'Quantity']
                    ).fillna(0)

                    pivot.columns = [f"{a}_{b}" for a, b in pivot.columns]
                    pivot = pivot.reset_index()

                    rename_dict = {
                        'avg_price_BUY': 'Buy_Avg_Price',
                        'avg_price_SELL': 'Sell_Avg_Price',
                        'Quantity_BUY': 'Buy_Quantity'
                    }
                    pivot = pivot.rename(columns={k: v for k, v in rename_dict.items() if k in pivot.columns})

                    if 'Quantity_SELL' in pivot.columns:
                        pivot = pivot.drop(columns=['Quantity_SELL'])

                    pivot['Hedge_Cost'] = (pivot.get('Sell_Avg_Price', 0) - pivot.get('Buy_Avg_Price', 0)) * pivot.get('Buy_Quantity', 0)

                    # User-wise summary
                    nifty = pivot[pivot['Symbol'].str.contains('NIFTY', case=False, na=False)] \
                            .groupby('User ID')['Hedge_Cost'].sum().reset_index(name='NIFTY_Hedge_Cost')
                    sensex = pivot[pivot['Symbol'].str.contains('SENSEX', case=False, na=False)] \
                             .groupby('User ID')['Hedge_Cost'].sum().reset_index(name='SENSEX_Hedge_Cost')

                    user_summary = pd.merge(nifty, sensex, on='User ID', how='outer').fillna(0)

                    # Save hedge files to bytes
                    hedge_buf = BytesIO()
                    pivot.to_csv(hedge_buf, index=False)
                    hedge_buf.seek(0)
                    st.session_state.tab2_hedge_bytes = hedge_buf.getvalue()

                    user_buf = BytesIO()
                    user_summary.to_csv(user_buf, index=False)
                    user_buf.seek(0)
                    st.session_state.tab2_user_hedge_bytes = user_buf.getvalue()

                    st.session_state.tab2_processed = True
                    st.success("Jainam processing finished successfully!")

                except Exception as e:
                    st.error(f"Error during Jainam processing: {str(e)}")
                    with st.expander("Show error details"):
                        st.exception(e)

    # ────────────────────────────────────────────────
    #  Download section – always visible after successful processing
    # ────────────────────────────────────────────────
    if st.session_state.tab2_processed:
        st.markdown("### Download Jainam Results")
        colA, colB, colC = st.columns(3)

        with colA:
            if st.session_state.tab2_pnl_bytes:
                st.download_button(
                    label="NF_SN pnl.xlsx",
                    data=st.session_state.tab2_pnl_bytes,
                    file_name="NF_SN_pnl.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_pnl_tab2"
                )

        with colB:
            if st.session_state.tab2_hedge_bytes:
                st.download_button(
                    label="hedge_cost_output.csv",
                    data=st.session_state.tab2_hedge_bytes,
                    file_name="hedge_cost_output.csv",
                    mime="text/csv",
                    key="dl_hedge_tab2"
                )

        with colC:
            if st.session_state.tab2_user_hedge_bytes:
                st.download_button(
                    label="user_wise_index_hedge_cost.csv",
                    data=st.session_state.tab2_user_hedge_bytes,
                    file_name="user_wise_index_hedge_cost.csv",
                    mime="text/csv",
                    key="dl_user_hedge_tab2"
                )

        st.info("Files are kept in memory until you refresh or restart the app. You can download multiple times.")
