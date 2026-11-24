import pandas as pd

# Načítanie hlavnej knihy
subor = "Vecné položky 07_2026.xlsx"

try:
    df = pd.read_excel(subor)
    print("--- Súbor bol úspešne načítaný! ---")
    print(f"Celkový počet riadkov: {len(df)}")
    print("\nZoznam stĺpcov v súbore:")
    print(df.columns.tolist())

except Exception as e:
    print(f"Chyba pri načítaní súboru: {e}")
    