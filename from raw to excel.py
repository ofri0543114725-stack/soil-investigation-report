def write_voc_sheet(ws, df, thresh_dict, t1col, t1lbl):
    df = df.copy()

    # (sample_id, depth) pairs: sample ASC, depth DESC
    pairs = sorted(
        df[["sample_id","depth"]].drop_duplicates().values.tolist(),
        key=lambda x: (sort_key(x[0]), -(x[1] or 0))
    )

    # ── HEADERS ───────────────────────────────────────────────────────────────
    # Row 1: A=קבוצה(merged A1:A2), B=קבוצה(merged B1:B2), C=שם התרכובת(merged C1:C2),
    #        D=CAS(merged D1:D2), E=VSL(merged E1:E2), F=TIER1(merged F1:F2),
    #        G=יחידות(merged G1:H2), I=שם קידוח row1 / עומק row2
    for ci, h in enumerate(["קבוצה","קבוצה","שם התרכובת","CAS","VSL",t1lbl],1):
        ws.merge_cells(start_row=1,start_column=ci,end_row=2,end_column=ci)
        style_hdr(ws.cell(1,ci,h), HDR_BLUE_FILL, sz=9)
        ws.cell(1,ci).alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
        ws.cell(1,ci).fill = HDR_BLUE_FILL
        ws.cell(1,ci).border = thin_border()

    # G1:H2 merged = יחידות
    ws.merge_cells(start_row=1,start_column=7,end_row=2,end_column=8)
    style_hdr(ws.cell(1,7,"יחידות"), HDR_BLUE_FILL, sz=9)
    ws.cell(1,7).alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.cell(1,7).fill = HDR_BLUE_FILL
    ws.cell(1,7).border = thin_border()

    # col I: שם קידוח (row1) / עומק (row2)
    style_hdr(ws.cell(1,9,"שם קידוח"), HDR_BLUE_FILL, sz=9)
    style_hdr(ws.cell(2,9,"עומק"),     HDR_BLUE_FILL, sz=9)

    # sample columns from col 10
    prev_sid = None
    sid_col_start = {}
    for ci, (sid, depth) in enumerate(pairs, 10):
        sid_val = sid if sid != prev_sid else None
        style_hdr(ws.cell(1,ci,sid_val), HDR_BLUE_FILL, sz=9)
        style_hdr(ws.cell(2,ci,depth),   HDR_BLUE_FILL, sz=9)
        if sid != prev_sid:
            sid_col_start[sid] = ci
        prev_sid = sid

    # merge row-1 sample name across depth columns
    for sid, sc in sid_col_start.items():
        cols = [ci for ci,(s,_) in enumerate(pairs,10) if s == sid]
        if len(cols) > 1:
            ws.merge_cells(start_row=1,start_column=sc,end_row=1,end_column=cols[-1])
            c = ws.cell(1,sc)
            c.alignment = Alignment(horizontal="center",vertical="center")
            c.fill = HDR_BLUE_FILL
            c.border = thin_border()

    # ALS data lookup: norm(compound) -> {(sid,depth): result_str}
    als_data = {}
    for _, r in df.iterrows():
        k = norm(r["compound"])
        if k not in als_data:
            als_data[k] = {}
        als_data[k][(r["sample_id"], r["depth"])] = r["result_str"]

    # add dot<->comma variants
    for k in list(als_data.keys()):
        for alt in (k.replace(".",","), k.replace(",", ".")):
            if alt not in als_data:
                als_data[alt] = als_data[k]

    # write data rows in exact VOC_COMPOUND_ORDER
    for row_i, (vs, grp, cmp) in enumerate(VOC_COMPOUND_ORDER, 3):
        vsl, tier1, cas = get_thresh(cmp, thresh_dict, t1col)
        cmp_key = norm(cmp)
        cmp_data = (
            als_data.get(cmp_key)
            or als_data.get(cmp_key.replace(".", ","))
            or {}
        )

        style_data(ws.cell(row_i,1, vs),    sz=9)
        style_data(ws.cell(row_i,2, grp),   sz=9)
        style_data(ws.cell(row_i,3, cmp),   sz=9)
        style_data(ws.cell(row_i,4, cas),   sz=9)
        style_data(ws.cell(row_i,5, vsl),   sz=9)
        style_data(ws.cell(row_i,6, tier1), sz=9)

        # G:H merged = units
        ws.merge_cells(start_row=row_i,start_column=7,end_row=row_i,end_column=8)
        style_data(ws.cell(row_i,7,"mg/kg"), sz=9)
        ws.cell(row_i,7).alignment = Alignment(horizontal="center",vertical="center")

        # col I empty
        style_data(ws.cell(row_i,9, None),  sz=9)

        # sample values from col 10
        for ci,(sid,depth) in enumerate(pairs,10):
            rs = cmp_data.get((sid, depth), "")
            style_data(ws.cell(row_i,ci,rs), check_exceed(rs, vsl, tier1), sz=9)

    # merge col A: VOCs / SVOCs
    for r1, r2, val in [(3,32,"VOCs"), (33,96,"SVOCs")]:
        ws.merge_cells(start_row=r1,start_column=1,end_row=r2,end_column=1)
        c = ws.cell(r1,1,val)
        c.font = Font(bold=True,name="Arial",size=9)
        c.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
        c.border = thin_border()

    # merge col B ranges (קבוצות)
    b_ranges = [
        (3, 12,  "Non-Halogenated VOCs"),
        (13,16,  "BTEX"),
        (17,32,  "Halogenated VOCs"),
        (33,37,  "Phenols & Naphtols"),
        (38,53,  "PAHs"),
        (54,57,  "Anilines"),
        (58,65,  "Aromatic Compounds"),
        (66,66,  "Alcohols"),
        (67,69,  "Aldehydes / Ketones"),
        (70,75,  "Chlorophenols"),
        (76,85,  "Nitroaromatic Compounds"),
        (86,88,  "Chlorinated Hydrocarbons"),
        (89,89,  "Nitrosoamines"),
        (90,90,  "Pesticides"),
        (91,96,  "Phthalates"),
    ]
    for r1, r2, val in b_ranges:
        if r2 > r1:
            ws.merge_cells(start_row=r1,start_column=2,end_row=r2,end_column=2)
        c = ws.cell(r1,2,val)
        c.font = Font(bold=True,name="Arial",size=9)
        c.alignment = Alignment(horizontal="center",vertical="center",wrap_text=True)
        c.border = thin_border()

    # column sizes + freeze
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 35
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 7
    ws.column_dimensions["H"].width = 7
    ws.column_dimensions["I"].width = 12
    for ci in range(10, 10+len(pairs)):
        ws.column_dimensions[get_column_letter(ci)].width = 10
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 15
    ws.freeze_panes = "J3"
