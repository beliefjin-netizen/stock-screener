# 1. 基本面 (營收年成長 > 0 AND 最新期毛利率 > 前一期毛利率 AND EPS > 0)
if use_fundamental:
  rev_col = next(
      (
          c
          for c in filtered_df.columns
          if "營收年成長" in c or "營收年增" in c or "營收" in c
      ),
      None,
  )
  eps_col = next(
      (c for c in filtered_df.columns if "EPS" in c or "每股盈餘" in c), None
  )

  # 精準辨識最新一期與前一期毛利率欄位
  gm_cols = [c for c in filtered_df.columns if "毛利率" in c]
  gm_curr_col = next((c for c in gm_cols if "前" not in c), None)
  gm_prev_col = next(
      (
          c
          for c in gm_cols
          if "前一季" in c or "前季" in c or "前1期" in c or "前期" in c
      ),
      None,
  )

  # 建立三項條件的布林遮罩 (Boolean Series)
  cond_rev = (
      (clean_number_series(filtered_df[rev_col]) > 0)
      if rev_col
      else pd.Series(True, index=filtered_df.index)
  )
  cond_eps = (
      (clean_number_series(filtered_df[eps_col]) > 0)
      if eps_col
      else pd.Series(True, index=filtered_df.index)
  )

  if gm_curr_col and gm_prev_col:
    gm_curr = clean_number_series(filtered_df[gm_curr_col])
    gm_prev = clean_number_series(filtered_df[gm_prev_col])
    cond_gm = (gm_curr > gm_prev) & gm_curr.notna() & gm_prev.notna()
  else:
    cond_gm = pd.Series(True, index=filtered_df.index)

  # 使用 & (AND) 複合邏輯，強制三者同時成立
  fundamental_mask = cond_rev & cond_gm & cond_eps
  filtered_df = filtered_df[fundamental_mask]