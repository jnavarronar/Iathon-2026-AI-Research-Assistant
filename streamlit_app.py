#!/usr/bin/env python3
import os
import io
import tempfile
from pathlib import Path
import streamlit as st

from AgentResearchAssitant_II import run_analysis


def streamlit_app():
    st.set_page_config(
        page_title="IAthon - AI Data Analysis",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    if "last_result" not in st.session_state:
        st.session_state.last_result = None

    st.markdown("""
        <style>
        .main-header {
            font-size: 3rem;
            font-weight: bold;
            text-align: center;
            color: #1E88E5;
            margin-bottom: 1rem;
        }
        .sub-header {
            text-align: center;
            color: #666;
            margin-bottom: 2rem;
        }
        .stProgress > div > div > div > div {
            background-color: #1E88E5;
        }
        .success-box {
            padding: 1rem;
            border-radius: 0.5rem;
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
            color: #155724;
        }
        .error-box {
            padding: 1rem;
            border-radius: 0.5rem;
            background-color: #f8d7da;
            border: 1px solid #f5c6cb;
            color: #721c24;
        }
        </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="main-header"> IAthon</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Automated AI-Powered Data Analysis</div>', unsafe_allow_html=True)

    with st.sidebar:
        st.header(" Configuration")

        api_key = st.text_input(
            "OpenAI API Key",
            type="password",
            value=os.getenv("OPENAI_API_KEY", ""),
            help="Enter your OpenAI API key or set OPENAI_API_KEY environment variable"
        )

        pexels_key = st.text_input(
            "PEXELS API Key (optional)",
            type="password",
            value=os.getenv("PEXELS_API_KEY", ""),
            help="Used to download a title image"
        )

        st.divider()

        st.subheader(" Analysis Settings")
        lang = st.selectbox(
            "Language",
            ["es", "en"],
            index=0,
            help="Report language"
        )

        output_format = st.selectbox(
            "Output Format",
            ["none", "pdf", "docx", "pptx"],
            index=0,
            help="Additional output format (requires Pandoc)"
        )

        st.divider()

        st.subheader(" Advanced Options")
        clean_figures = st.checkbox(
            "Clean figures directory",
            value=False,
            help="Delete existing figures before analysis"
        )

        verbose = st.checkbox(
            "Verbose output",
            value=False,
            help="Show detailed execution logs"
        )

        auto_fix = st.checkbox(
            "Auto-fix data issues",
            value=True,
            help="Automatically fix common data quality issues"
        )

        st.divider()

        st.subheader(" About")
        st.info("""
        **IAthon** uses AI to automatically:
        - Validate data quality
        - Generate visualizations
        - Perform statistical tests
        - Create comprehensive reports

        Powered by OpenAI GPT-4o-mini
        """)

    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader(" Dataset Upload")
        upload = st.file_uploader(
            "Upload your dataset",
            type=["csv", "xlsx", "xls"],
            help="Supports CSV and Excel files"
        )

        if upload:
            st.success(f" File uploaded: **{upload.name}** ({upload.size / 1024:.1f} KB)")

    with col2:
        st.subheader(" Output Settings")
        output_name = st.text_input(
            "Report filename",
            value="report.md",
            help="Name for the output report"
        )

    st.divider()

    with st.expander(" Custom Analysis Prompt (Optional)", expanded=False):
        st.markdown("""
        Add specific instructions for the analysis:
        - Domain context
        - Specific questions to answer
        - Analysis focus areas
        """)
        custom_prompt = st.text_area(
            "Custom prompt",
            height=150,
            placeholder="Example: This dataset contains customer purchase data. Focus on identifying spending patterns by demographic groups...",
            label_visibility="collapsed"
        )

    st.divider()

    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        run_button = st.button(" Run Analysis", use_container_width=True, type="primary")

    if run_button:
        if not upload:
            st.error(" Please upload a dataset first.")
            return

        if not api_key:
            st.error(" Please provide an OpenAI API key in the sidebar.")
            return

        if not output_name.lower().endswith(".md"):
            output_name = f"{output_name}.md"

        output_dir = Path("results")
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / output_name

        suffix = Path(upload.name).suffix
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(upload.getbuffer())
            input_path = Path(tmp.name)

        progress_bar = st.progress(0)
        status_text = st.empty()
        details_text = st.empty()

        def update_progress(step, max_steps, status, details):
            progress = int((step / max_steps) * 100)
            progress_bar.progress(progress)
            status_text.markdown(f"**Status:** {status}")
            details_text.markdown(f"_{details}_")

        try:
            if pexels_key:
                os.environ["PEXELS_API_KEY"] = pexels_key
            with st.spinner(" Running analysis..."):
                result = run_analysis(
                    input_path=input_path,
                    output_path=output_path,
                    output_format=None if output_format == "none" else output_format,
                    api_key=api_key,
                    verbose=verbose,
                    custom_prompt=custom_prompt.strip() if custom_prompt.strip() else None,
                    prompt_file=None,
                    lang=lang,
                    clean_figures=clean_figures,
                    auto_fix=auto_fix,
                    progress_callback=update_progress,
                )

            progress_bar.progress(100)
            status_text.markdown("**Status:** Complete")
            details_text.markdown("_Analysis finished successfully_")

            st.session_state.last_result = result
            st.balloons()
            st.markdown('<div class="success-box"> <b>Analysis completed successfully!</b></div>', unsafe_allow_html=True)

        except Exception as e:
            st.markdown(f'<div class="error-box"> <b>Analysis failed:</b> {str(e)}</div>', unsafe_allow_html=True)
            st.exception(e)

        finally:
            try:
                input_path.unlink()
            except Exception:
                pass

    if st.session_state.last_result:
        result = st.session_state.last_result
        st.divider()
        st.header(" Results")

        tab1, tab2, tab3, tab4 = st.tabs([" Report", " Figures", " Cost Summary", " Downloads"])

        with tab1:
            st.subheader("Generated Report")
            report_text = result["report_text"]
            st.markdown(report_text)

        with tab2:
            st.subheader("Generated Visualizations")
            fig_dir = result["figures_dir"]
            pngs = sorted(fig_dir.glob("*.png"))
            jpgs = sorted(fig_dir.glob("*.jpg"))
            figs = (result["new_figs"] if "new_figs" in result else pngs) + [
                p for p in jpgs if p.name == "title_image.jpg"
            ]

            if figs:
                cols_per_row = 2
                for i in range(0, len(figs), cols_per_row):
                    cols = st.columns(cols_per_row)
                    for j, col in enumerate(cols):
                        if i + j < len(figs):
                            fig = figs[i + j]
                            with col:
                                st.image(str(fig), caption=fig.name, use_container_width=True)
            else:
                st.info("No figures generated")

        with tab3:
            st.subheader("Token Usage & Cost")
            log_path = result["report_md_path"].parent / "token_usage.log"
            if log_path.exists():
                log_content = log_path.read_text()
                st.code(log_content, language="text")
            else:
                st.info("Cost summary not available")

        with tab4:
            st.subheader("Download Files")
            col_d1, col_d2 = st.columns(2)

            with col_d1:
                report_path = result["report_md_path"]
                st.download_button(
                    label=" Download Report (Markdown)",
                    data=report_path.read_text(encoding='utf-8'),
                    file_name=report_path.name,
                    mime="text/markdown",
                    use_container_width=True,
                )

            with col_d2:
                if result["out_file"] and result["out_file"].exists():
                    st.download_button(
                        label=f" Download {result['out_file'].suffix.upper()} Report",
                        data=result['out_file'].read_bytes(),
                        file_name=result['out_file'].name,
                        use_container_width=True,
                    )

            if figs:
                import zipfile
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for fig in figs:
                        zip_file.write(fig, fig.name)

                st.download_button(
                    label=" Download All Figures (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="figures.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

if __name__ == "__main__":
    streamlit_app()