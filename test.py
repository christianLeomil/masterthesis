joules_per_cm_squared = 1000  # Example value in J/cm²

joules_per_m_squared = joules_per_cm_squared * 10000
kwh_per_m_squared = joules_per_m_squared * 0.00000027778

print(f"{joules_per_cm_squared} J/cm² is approximately {kwh_per_m_squared} kWh/m².")
