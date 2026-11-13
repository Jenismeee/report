        background: rgba(255,255,255,0.85);
        color: #526579;
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: 0.04em;
        margin-bottom: 1rem;
    }
    .landing h1 {
        font-size: 2.8rem;
        line-height: 1.15;
        letter-spacing: -0.04em;
        margin: 0 0 0.8rem 0;
        color: #17212b;
    }
    .landing p {
        max-width: 650px;
        margin: 0 auto;
        color: #687787;
        font-size: 1.03rem;
        line-height: 1.7;
    }
    [data-testid="stFileUploader"] {
        max-width: 760px;
        margin: 1.2rem auto 0 auto;
    }
    [data-testid="stFileUploaderDropzone"] {
        border: 1.5px dashed #b9c8d8 !important;
        border-radius: 16px !important;
        background: rgba(255,255,255,0.9) !important;
        padding: 1.25rem 1rem !important;
    }
    div.stButton {
        max-width: 760px;
        margin: 0.9rem auto 0 auto;
    }
    div.stButton > button {
        width: 100%;
        min-height: 2.8rem;
        border-radius: 10px;
        font-weight: 600;
        border: 0;
    }
    </style>
    <div class="landing">
        <div class="landing-badge">CARGO OPERATIONS · AUTOMATED REPORTING</div>
        <h1>📦 Cargo Operations Dashboard</h1>
        <p>
            Turn structured cargo data into clear operations and declarant reports
            in a few clicks. Upload one Excel file to get started.
        </p>
    </div>
    """,
    unsafe_allow_html=True
)

file = st.file_uploader(
    'Upload Excel File',
    type=['xlsx'],
    help='Supported format: .xlsx'
)

if file:
    if st.button('Generate Reports', type='primary'):
        try:
            file_bytes = file.getvalue()
            df = read_overall(file_bytes)
            road_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='Road')
            export_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='Export')

            # Existing Cargo Operations Report
            cargo_result = analyze(df)
            cargo_months = cargo_result[-1]
            raw_coe, recognized_coe = coe_validation(df)
            cargo_report = build_html(df)

            # New Declarant Monthly Report
            declarant_report = build_declarant_html(df, road_df, export_df)
            declarant_result = analyze_declarant(df, road_df, export_df)
            declarant_months = declarant_result[2]

            if raw_coe and recognized_coe == 0:
                st.error(
                    f"Cainiao-COE check failed: the Excel contains {raw_coe} COE row(s), but 0 were recognized."
                )
            else:
                st.success(
                    f"Reports generated successfully. Cargo: {len(cargo_months)} month(s); "
                    f"Declarant: {len(declarant_months)} month(s) from Jan 2026 onward; "
                    f"Cainiao-COE rows recognized: {recognized_coe}."
                )

            tab1, tab2 = st.tabs([
                '📦 Cargo Operations Report',
                '🛃 Declarant Monthly Report'
            ])

            with tab1:
                st.components.v1.html(
                    cargo_report,
                    height=1200,
                    scrolling=True
                )
                st.download_button(
                    '⬇️ Download Cargo Operations HTML',
                    data=cargo_report.encode('utf-8'),
                    file_name='Cargo_Operations_Dashboard.html',
                    mime='text/html',
                    key='download_cargo'
                )

            with tab2:
                st.components.v1.html(
                    declarant_report,
                    height=1200,
                    scrolling=True
                )
                st.download_button(
                    '⬇️ Download Declarant Monthly HTML',
                    data=declarant_report.encode('utf-8'),
                    file_name='Declarant_Monthly_Report.html',
                    mime='text/html',
                    key='download_declarant'
                )

        except Exception as e:
            st.error(f'Processing failed: {e}')
