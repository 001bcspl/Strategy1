import pandas as pd

FILE_PATH = 'edata_b.xlsx'
OUTPUT_FILE = 'logs/stg1_edata_b.xlsx'
AMOUNT_INVESTED = 5000
EACH_BUNDLE = 35


df = pd.read_excel(FILE_PATH)
df.columns = df.columns.str.strip().str.upper()
df.rename(columns={'LTPCE': 'CE', 'LTPPE': 'PE'}, inplace=True)


log_rows = []
ce_prev = df.iloc[0]['CE']
pe_prev = df.iloc[0]['PE']
crossover_summary = {}
i = 1

while i < len(df):
    row = df.iloc[i]
    ce_ltp = row['CE']
    pe_ltp = row['PE']
    time = row['TIME']

    if ce_prev < pe_prev and ce_ltp > pe_ltp:
      
        crossover = round(ce_ltp, 2)
        sl = round(crossover * 0.7, 2)
        sl_value = round(crossover - sl, 2)

        if sl_value == 0:
            log_rows.append([time.strftime("%H:%M:%S"), "SL=0_ERR", "N", "N", "N", "N"])
            i += 1
            continue

        target = round(crossover + 2 * sl_value, 2)
        total_bundles = round(AMOUNT_INVESTED / (EACH_BUNDLE * sl_value))
        loss_on_sl = round((sl - crossover) * EACH_BUNDLE * total_bundles, 2)
        profit_on_target = round((target - crossover) * EACH_BUNDLE * total_bundles, 2)

        
        crossover_summary = {
            'time': time.strftime('%H:%M:%S'),
            'crossover': crossover,
            'sl': sl,
            'target': target,
            'sl_value': sl_value,
            'total_bundles': total_bundles,
            'loss_on_sl': loss_on_sl,
            'profit_on_target': profit_on_target,
            'sl_hit_time': None,
            'sl_hit_ce': None,
            'target_hit_time': None
        }

     
        log_rows.append([time.strftime("%H:%M:%S"), "CROSSOVER", crossover, sl, total_bundles, loss_on_sl])

        j = i + 1
        while j < len(df):
            next_row = df.iloc[j]
            next_ce = next_row['CE']
            next_pe = next_row['PE']
            next_time = next_row['TIME']

            if next_ce >= target:
                crossover_summary['target_hit_time'] = next_time.strftime('%H:%M:%S')
                break

            elif next_ce <= sl:
                crossover_summary['sl_hit_time'] = next_time.strftime('%H:%M:%S')
                crossover_summary['sl_hit_ce'] = next_ce

                
                log_rows.append([
                    next_time.strftime("%H:%M:%S"),
                    "SL_HIT",
                    crossover,
                    sl,
                    total_bundles,
                    loss_on_sl
                ])
                break

            
            log_rows.append([next_time.strftime("%H:%M:%S"), "NONE", "N", "N", "N", "N"])
            j += 1
        i = j
    else:
        log_rows.append([time.strftime("%H:%M:%S"), "NONE", "N", "N", "N", "N"])
        ce_prev = ce_ltp
        pe_prev = pe_ltp
        i += 1


print("\n--- Log Table ---")
print("{:<10} {:<11} {:<8} {:<8} {:<6} {:<10}".format("TIME", "EVENT", "BP", "SL", "QTY", "PL"))
print("-" * 60)
for row in log_rows:
    print("{:<10} {:<11} {:<8} {:<8} {:<6} {:<10}".format(*row))


if crossover_summary:
    print("\n--- Crossover Summary ---")
    print(f"Cross Over       : {crossover_summary['crossover']}")
    print(f"SL               : {crossover_summary['sl']}")
    print(f"Target           : {crossover_summary['target']}")
    print(f"Amount Invested  : {AMOUNT_INVESTED}")
    print(f"Each Bundle      : {EACH_BUNDLE}")
    print(f"Total Bundles    : {crossover_summary['total_bundles']}")
    print(f"SL Value         : {crossover_summary['sl_value']}")
    print(f"Loss on SL Hit   : {crossover_summary['loss_on_sl']}")
    print(f"Profit on Target : {crossover_summary['profit_on_target']}")

    if crossover_summary['sl_hit_time']:
        print(f"→ SL hit at {crossover_summary['sl_hit_time']}, CE: {crossover_summary['sl_hit_ce']}")
    elif crossover_summary['target_hit_time']:
        print(f"→ Target hit at {crossover_summary['target_hit_time']}, CE >= {crossover_summary['target']}")
    else:
        print("→ Neither SL nor Target hit by end of data.")

output_df = pd.DataFrame(log_rows, columns=['TIME', 'EVENT', 'BP', 'SL', 'QTY', 'PL'])
output_df.to_excel(OUTPUT_FILE, index=False)
