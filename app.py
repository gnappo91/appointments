import streamlit as st
import pandas as pd
from io import BytesIO
from pdb import set_trace

# ------------------------
# Core scheduling functions
# ------------------------
def try_assign(patient_name, avail_slots, agenda_df, start_hour, end_hour):
    from datetime import time

    for slot in avail_slots:
        if pd.isna(slot):
            continue
        try:
            slot = pd.to_datetime(slot)
        except Exception:
            continue

        # (optional) make 19:00 inclusive; if you want exclusive, use < time(end_hour,0)
        in_window = time(start_hour, 0) <= slot.time() <= time(end_hour, 0)
        if not in_window:
            continue

        # Match date column
        parsed = pd.to_datetime(agenda_df.columns[1:], format="%d/%m/%y", errors="coerce")
        mask = parsed.date == slot.date()
        if not mask.any():
            continue
        date_col = agenda_df.columns[1:][mask][0]

        # Match time row
        agenda_df["Orario"] = agenda_df["Orario"].astype(str).str[:5]
        slot_time = slot.strftime("%H:%M")
        if slot_time not in set(agenda_df["Orario"]):
            continue

        row_idx = agenda_df["Orario"].eq(slot_time).idxmax()

        if pd.isna(agenda_df.at[row_idx, date_col]) or agenda_df.at[row_idx, date_col] == "":
            agenda_df.at[row_idx, date_col] = patient_name
            return True

    return False


def assign_slots(agenda_df, disp_df):
    non_assigned = []

    for _, row in disp_df.iterrows():
        patient = row.iloc[0]
        avail_slots = row.iloc[1:6].values

        # Preferred 17:00-19:00
        assigned = try_assign(patient, avail_slots, agenda_df, 17, 19)

        # If not assigned, try 7:00-22:00
        if not assigned:
            assigned = try_assign(patient, avail_slots, agenda_df, 7, 22)

        if not assigned:
            non_assigned.append(patient)

    return agenda_df, non_assigned


# ------------------------
# Streamlit UI
# ------------------------
st.title("📅 Organizzazione appuntamenti")

uploaded_file = st.file_uploader(
    "Carica file excel con 2 fogli: 'Calendario' e 'Disponibilità'",
    type=["xls", "xlsx"]
    )

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    if not set(["Calendario", "Disponibilità"]).issubset(xls.sheet_names):
        st.error("Excel must contain 'Calendario' and 'Disponibilità' sheets.")
    else:
        agenda_df = pd.read_excel(xls, sheet_name="Calendario")
        disp_df = pd.read_excel(xls, sheet_name="Disponibilità")
        disp_df = disp_df.loc[:,(disp_df.isna().all(axis=0)==False).values]
        st.subheader("Disponibilità")
        st.dataframe(disp_df)

        if st.button("Assegna orari"):
            updated_agenda, non_assigned = assign_slots(agenda_df.copy(), disp_df)

            st.success("✅ Assegnazione completata!")

            st.subheader("Calendario aggiornato")
            st.dataframe(updated_agenda)

            if non_assigned:
                st.subheader("Non Assegnati")
                st.write("\n".join(f"- {item}" for item in non_assigned))

            # Prepare Excel download
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                updated_agenda.to_excel(writer, sheet_name="Calendario", index=False)
                disp_df.to_excel(writer, sheet_name="Disponibilità", index=False)
                pd.DataFrame(non_assigned, columns=["Pazienti non assegnati"]).to_excel(writer, sheet_name="Non Assegnati", index=False)

            st.download_button(
                label="📥 Scarica l'excel con gli appuntamenti assegnati",
                data=output.getvalue(),
                file_name="updated_agenda.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )