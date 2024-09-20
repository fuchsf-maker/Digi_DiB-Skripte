import pandas as pd
import streamlit as st
import io

# Funktion zur Umwandlung von DataFrame in eine Arrow-kompatible Version
def make_arrow_compatible(df):
    # Konvertieren Sie alle Spalten in String, um sicherzustellen, dass sie Arrow-kompatibel sind
    for col in df.columns:
        df[col] = df[col].astype(str)
    return df

# Titel der Streamlit-Anwendung
st.title("Deutsche Nationalbibliothek: Dublettencheck")

# Sidebar für Navigation
st.sidebar.title("Dublettencheck...")
page = st.sidebar.radio("Wählen Sie einen Bereich", [
    "Exakte Werte",
    "Börsenblatt"
])

# Beschreibung des Dublettenchecks
if page == "Exakte Werte":
    st.write(
        "Diese Applikation wertet auf Basis von Exceltabellen bestimmte Kategorien aus. Zum Start bitte die Excelliste mit den Kategorien, die untersucht werden sollen mittels WinIBW oder PICA-RS erstellen und anschließend zur Analyse im Tool hochladen. In einem 2. Tabellenblatt wird, falls gegeben, die Duplikate mit URN ausgegeben."
    )

    st.subheader("Auswertung nach exakten Werten")

    st.write(
        "Zur Importüberprüfung. Die Inhalte in der gewählten Kategorie werden zeichengenau überprüft."
    )

    # Datei-Upload
    uploaded_file = st.file_uploader("Laden Sie eine Excel-Datei hoch.", type=["xlsx"])

    # Überprüfen, ob eine Datei hochgeladen wurde
    if uploaded_file:
        try:
            # Versuche, die Excel-Datei mit dem 'openpyxl'-Engine zu laden
            df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            # Datentypen überprüfen und ggf. bereinigen
            st.write("Datentypen der geladenen Daten:")
            st.write(df.dtypes)
            
            # Umwandlung des DataFrames in eine Arrow-kompatible Version
            df = make_arrow_compatible(df)
            
            # Überprüfen, ob DataFrame geladen wurde
            st.write("Daten erfolgreich geladen:")
            st.write(df.head())
            
            # Anzeigen der Spaltennamen zur Auswahl
            columns = df.columns.tolist()
            selected_column = st.selectbox("Wählen Sie die zu überprüfende Spalte", columns)
            
            # Los-Button zum Starten der Analyse
            if st.button("Dubletten überprüfen"):
                # Funktion zur Überprüfung von Duplikaten innerhalb einer Gruppe
                def check_duplicates(group):
                    return group[group.duplicated(subset=selected_column, keep=False)]
                
                # Gruppieren nach 'Überordnung' und Überprüfung auf Duplikate innerhalb jeder Gruppe
                duplicates = df.groupby('Überordnung').apply(check_duplicates).reset_index(drop=True)
                
                # Überprüfen, ob Duplikate gefunden wurden
                if not duplicates.empty:
                    st.write("Folgende Duplikate wurden gefunden:")
                    st.dataframe(duplicates)
                    
                    # Möglichkeit, die Duplikate herunterzuladen
                    def convert_df(df):
                        towrite = io.BytesIO()
                        with pd.ExcelWriter(towrite, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False, sheet_name='Duplicates')
                            # Überprüfen auf Einträge ohne URN
                            if 'URN' in df.columns:
                                entries_without_urn = df[df['URN'].isnull()]
                                if not entries_without_urn.empty:
                                    entries_without_urn.to_excel(writer, sheet_name='Duplicates_without_URN', index=False)
                        return towrite.getvalue()
                    
                    excel_file = convert_df(duplicates)
                    st.download_button(
                        label="Duplikate als Excel-Datei herunterladen",
                        data=excel_file,
                        file_name='duplicates.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                else:
                    st.write("Keine Duplikate gefunden.")
        except Exception as e:
            st.error(f"Beim Verarbeiten der Datei ist ein Fehler aufgetreten: {e}")

elif page == "Börsenblatt":
    st.subheader("Börsenblatt")
    st.write("""
              Hier wird der Dublettencheck für das Börsenblatt beschrieben und implementiert. Die Jahrgänge werden abgeglichen. Anschließend wird geprüft, in welchem Erscheinungsjahr die Daten veröffentlicht wurden, um herauszufinden, wo die Dubletten liegen, da es identische Lieferungen in verschiedenen Jahren gab.

    Folgende Felder werden in folgender Benennung benötigt:

    Satzart:                 002@ $0
             
    Jahr:                  011@ $a
             
    Überordnung:          021A $9
             
    Digicode:                 017C $a
             
    Jahrgang:          021B $l
             
    URN:                   K004U $0
             
             """)
    
    # Datei-Upload
    uploaded_file = st.file_uploader("Laden Sie eine Excel-Datei hoch.", type=["xlsx"])

    # Überprüfen, ob eine Datei hochgeladen wurde
    if uploaded_file:
        try:
            # Versuche, die Excel-Datei mit dem 'openpyxl'-Engine zu laden
            df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            # Datentypen überprüfen und ggf. bereinigen
            st.write("Datentypen der geladenen Daten:")
            st.write(df.dtypes)
            
            # Umwandlung des DataFrames in eine Arrow-kompatible Version
            df = make_arrow_compatible(df)
            
            # Überprüfen, ob DataFrame geladen wurde
            st.write("Daten erfolgreich geladen:")
            st.write(df.head())

            # Filtern nach Digicode-Wert 'd034'
            df_filtered = df[df['Digicode'] == 'd034']
            
            # Funktion zur Überprüfung von Duplikaten innerhalb einer Gruppe
            def check_duplicates(group):
                duplicates = group[group.duplicated(subset=['Jahrgang', 'Erscheinungsjahr'], keep=False)]
                return duplicates

            # Gruppieren nach 'Ueberordnung' und Überprüfung auf Duplikate innerhalb jeder Gruppe
            duplicates = df_filtered.groupby('Ueberordnung').apply(check_duplicates).reset_index(drop=True)
            
            # Wenn es Duplikate gibt, diese in einem neuen DataFrame speichern
            if not duplicates.empty:
                st.write("Folgende Duplikate wurden gefunden:")
                st.dataframe(duplicates)

                def convert_df(df):
                    towrite = io.BytesIO()
                    with pd.ExcelWriter(towrite, engine='openpyxl') as writer:
                        for ueberordnung, group in df.groupby('Ueberordnung'):
                            group.to_excel(writer, index=False, sheet_name=str(ueberordnung)[:31])
                        # Überprüfen auf Einträge ohne URN
                        entries_without_urn = df[df['URN'].isnull()]
                        if not entries_without_urn.empty:
                            entries_without_urn.to_excel(writer, sheet_name='Duplicates_without_URN', index=False)
                    return towrite.getvalue()
                
                excel_file = convert_df(duplicates)
                st.download_button(
                    label="Duplikate als Excel-Datei herunterladen",
                    data=excel_file,
                    file_name='duplicates_year.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.write("Keine Duplikate gefunden.")
        except Exception as e:
            st.error(f"Beim Verarbeiten der Datei ist ein Fehler aufgetreten: {e}")

st.write(" ")
st.write(" ")
st.markdown(" ###### Bitte bei inkorrekten Auswertungen, die aufgrund von Analysefehlern auftreten an f.fuchs@dnb.de melden. ")
st.write(" ")
