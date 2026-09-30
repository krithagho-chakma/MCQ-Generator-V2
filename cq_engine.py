import streamlit as st

# We added the two arguments here so app.py doesn't crash when trying to pass them
def run_cq_interface(model_choice, api_key_input):
    """Placeholder interface for the upcoming CQ Engine."""
    st.info(
        "**🚧 Under Construction!**\n\n"
        "The Creative Questions module is currently in development. "
        "Soon, you will be able to generate and format full structured creative questions "
        "just like the MCQ engine. Stay tuned!", 
        icon="⏳"
    )
    
    # A disabled button to visually indicate it's a future feature
    st.button("🚀 Generate CQs", disabled=True, use_container_width=True)
