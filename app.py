"""Streamlit screen for the inventory bot.  Run:  streamlit run app.py"""
import contextlib
import glob
import io
import os
import streamlit as st

st.set_page_config(page_title="Warehouse Assistant", page_icon="📦", layout="wide")

# ---------------------------------------------------------------- load bot.py
try:
    import bot                       # your bot.py: loads Excel and builds the tables (runs once)
except FileNotFoundError as e:
    st.error("Could not find the Excel file. Keep 'Inventory Data.xlsx' in the same folder as app.py.\n\n" + str(e))
    st.stop()

# ---------------------------------------------------------------- email: confirm with a button
# bot.py asks for "yes" in the terminal. Here we stage the draft, show it on screen,
# and only send when the person presses the Send button.
if not hasattr(bot, "_orig_send"):
    bot._orig_send = bot.tool_send_email

def staged_send(to, subject, body, attach_file=None):
    bot.ask_human = lambda prompt: "__not_yet__"            # reaching the question means all checks passed
    with contextlib.redirect_stdout(io.StringIO()):
        res = bot._orig_send(to, subject, body, attach_file)
    if res.get("sent") is False and "declined" in str(res.get("note", "")):
        st.session_state.pending_email = {"to": to, "subject": subject, "body": body, "attach_file": attach_file}
        return {"sent": False, "note": "Draft prepared and shown to the user. It is NOT sent until the user presses Send in the app."}
    return res                                              # an error from the safety checks

bot.TOOL_FUNCS["send_email"] = staged_send

def send_pending():
    p = st.session_state.pop("pending_email")
    bot.ask_human = lambda prompt: "yes"                    # the person just pressed the button
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            res = bot._orig_send(p["to"], p["subject"], p["body"], p["attach_file"])
    except Exception as e:                                  # wrong password, no internet ...
        res = {"error": str(e)}
    note = ("✅ Email sent to " + p["to"]) if res.get("sent") else "❌ Not sent: " + str(res.get("error") or res.get("note"))
    st.session_state.messages.append({"role": "assistant", "content": note, "files": []})
    st.session_state.history.append({"role": "assistant", "content": note})

def cancel_pending():
    st.session_state.pop("pending_email", None)
    st.session_state.messages.append({"role": "assistant", "content": "Email cancelled. Nothing was sent.", "files": []})
    st.session_state.history.append({"role": "assistant", "content": "Email cancelled. Nothing was sent."})

# ---------------------------------------------------------------- session state
st.session_state.setdefault("messages", [])      # what is drawn on screen
st.session_state.setdefault("history", [])       # text-only memory sent to the AI
ai_mode = bool(os.environ.get("ANTHROPIC_API_KEY"))
t = bot.t

def mime(path):
    return "application/pdf" if path.lower().endswith(".pdf") else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

def answer_question(q):
    if not ai_mode:
        return bot.handle(q)
    try:
        return bot.ask_llm(q, st.session_state.history[-6:])
    except Exception as e:
        return "AI error: " + str(e)

# ---------------------------------------------------------------- sidebar
with st.sidebar:
    st.header("📦 Warehouse Assistant")
    if ai_mode:
        st.success("AI mode: ask anything")
    else:
        st.info("Keyword mode (no API key). AI questions are switched off.")

    if not ai_mode:
        with st.expander("Add an API key (this computer only)"):
            key = st.text_input("Anthropic API key", type="password")
            if st.button("Use key") and key.strip():
                os.environ["ANTHROPIC_API_KEY"] = key.strip()
                st.rerun()

    st.subheader("Quick questions")
    quick = None
    for label, q, needs_ai in [
        ("What's low?", "what's low", False),
        ("Top 5 fast moving", "top 5 fast moving", False),
        ("Overstock", "overstock", False),
        ("Dead stock", "dead stock", False),
        ("Summary", "summary", False),
        ("🤖 Forecast total demand", "Forecast total demand for next month", True),
        ("🤖 What if lead time doubles?", "What if lead time doubles for all items?", True),
        ("🤖 Make a PDF reorder list", "Make a PDF of the reorder list", True),
        ("🤖 Which supplier is slowest?", "Which supplier is slowest?", True),
    ]:
        if st.button(label, width="stretch", disabled=needs_ai and not ai_mode):
            quick = q

    with st.expander("📋 Record a stock count"):
        c_sku = st.text_input("SKU", placeholder="1193BA")
        c_qty = st.number_input("Counted quantity (Kg)", min_value=0.0, step=1.0)
        if st.button("Save count"):
            r = bot.tool_record_count(c_sku, c_qty)
            if "error" in r:
                st.error(r["error"])
            else:
                st.success(f"{r['sku']}: system {r['system_qty']:,.0f} Kg, counted {r['counted_qty']:,.0f} Kg, variance {r['variance']:+,.0f} Kg")

    if st.button("Clear chat", width="stretch"):
        st.session_state.messages, st.session_state.history = [], []
        st.session_state.pop("pending_email", None)
        st.rerun()

# ---------------------------------------------------------------- header numbers
st.title("📦 Warehouse Assistant")
c1, c2, c3, c4 = st.columns(4)
c1.metric("SKUs", f"{len(t)}")
c2.metric("Need reorder now", int(t["status"].isin(["STOCKOUT", "CRITICAL"]).sum()))
c3.metric("Overstock", int((t["status"] == "OVERSTOCK").sum()))
c4.metric("Stock value", f"{t['value'].sum() / 1e6:,.1f} M")

tab_chat, tab_dash = st.tabs(["💬 Chat", "📊 Dashboard"])

# ---------------------------------------------------------------- chat tab
with tab_chat:
    if not st.session_state.messages:
        st.caption("Try: `price of 1193BA`, `what's low`, or use the buttons on the left. Anything outside the data gets: Not updated.")

    for i, m in enumerate(st.session_state.messages):
        with st.chat_message(m["role"]):
            st.markdown(m["content"])
            for f in m.get("files", []):
                if os.path.exists(f):
                    with open(f, "rb") as fh:
                        st.download_button("⬇ " + os.path.basename(f), fh.read(), file_name=os.path.basename(f),
                                           mime=mime(f), key=f"dl_{i}_{os.path.basename(f)}")

    prompt = st.chat_input("Ask about stock, price, lead time, orders...") or quick
    if prompt:
        before = set(glob.glob(os.path.join(bot.REPORT_DIR, "*")))
        st.session_state.messages.append({"role": "user", "content": prompt, "files": []})
        with st.chat_message("user"):
            st.markdown(prompt)
        with st.chat_message("assistant"):
            with st.spinner("Thinking..."):
                answer = answer_question(prompt)
            st.markdown(answer)
            new_files = sorted(set(glob.glob(os.path.join(bot.REPORT_DIR, "*"))) - before)
            for f in new_files:
                with open(f, "rb") as fh:
                    st.download_button("⬇ " + os.path.basename(f), fh.read(), file_name=os.path.basename(f),
                                       mime=mime(f), key=f"dl_new_{os.path.basename(f)}")
        st.session_state.messages.append({"role": "assistant", "content": answer, "files": new_files})
        st.session_state.history += [{"role": "user", "content": prompt}, {"role": "assistant", "content": answer}]

    # email draft card: nothing is sent until the person presses Send
    if "pending_email" in st.session_state:
        p = st.session_state.pending_email
        with st.container(border=True):
            st.subheader("✉️ Email draft: not sent yet")
            st.write(f"**To:** {p['to']}")
            st.write(f"**Subject:** {p['subject']}")
            st.write(f"**Attachment:** {os.path.basename(p['attach_file']) if p['attach_file'] else 'none'}")
            st.text(p["body"])
            b1, b2, _ = st.columns([1, 1, 4])
            b1.button("Send email", type="primary", on_click=send_pending)
            b2.button("Cancel", on_click=cancel_pending)

# ---------------------------------------------------------------- dashboard tab
with tab_dash:
    left, right = st.columns([1, 2])
    with left:
        st.subheader("SKUs by status")
        st.bar_chart(t["status"].value_counts())
    with right:
        st.subheader("Most urgent SKUs")
        urgent = bot.restock_list(50)[["SKU ID", "status", "abc", "stock", "cover", "lead_avg", "lead_max", "order_qty"]].round(1)
        st.dataframe(urgent, width="stretch", hide_index=True)

    st.subheader("Download a report")
    d1, d2, d3 = st.columns(3)
    for col, rep, fmt, label in [(d1, "reorder", "excel", "Reorder list (Excel)"),
                                 (d2, "reorder", "pdf", "Reorder list (PDF)"),
                                 (d3, "overstock", "excel", "Overstock list (Excel)")]:
        if col.button("Prepare: " + label):
            r = bot.tool_export_report(rep, fmt)
            if "error" in r:
                col.error(r["error"])
            else:
                st.session_state["dash_file"] = r["file"]
    f = st.session_state.get("dash_file")
    if f and os.path.exists(f):
        with open(f, "rb") as fh:
            st.download_button("⬇ Download " + os.path.basename(f), fh.read(), file_name=os.path.basename(f), mime=mime(f))