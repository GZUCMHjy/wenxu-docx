"""文序 v0.6 entry point; v0.5 remains available for historical regression."""
import streamlit as st

if st.query_params.get("legacy") == "1":
    import runpy
    from pathlib import Path
    runpy.run_path(str(Path(__file__).with_name("legacy_app.py")), run_name="__main__")
else:
    from v06_ui import main
    main()
