import streamlit as st
import pandas as pd
import re
import matplotlib.pyplot as plt
from io import BytesIO
from docx import Document

# List of punctuation keys to analyze
PUNCTUATION_KEYS = [
    "apostrophes", "colons", "commas", "curly_brackets", "double_inverted_commas",
    "ellipses", "em_dashes", "en_dashes", "exclamation_marks", "full_stops",
    "hyphens", "other_punctuation_marks", "question_marks", "round_brackets",
    "semicolons", "slashes", "square_brackets", "vertical_bars"
]

# Function to extract all text from a Word document
def extract_text(file):
    doc = Document(file)
    content = ""

    for p in doc.paragraphs:
        content += p.text + " "

    for section in doc.sections:
        if section.header:
            for p in section.header.paragraphs:
                content += p.text + " "
        if section.footer:
            for p in section.footer.paragraphs:
                content += p.text + " "

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                content += cell.text + " "

    content = re.sub(r'(\s?\.\s?){2,}', '...', content)
    return content

# Count punctuation marks
def count_punctuation(content):
    return {
        "apostrophes": len(re.findall(r"[\'’]", content)),
        "colons": len(re.findall(r":", content)),
        "commas": len(re.findall(r",", content)),
        "curly_brackets": len(re.findall(r"\{|\}", content)),
        "double_inverted_commas": len(re.findall(r"[“”\"]", content)),
        "ellipses": len(re.findall(r"…", content)) + content.count("..."),
        "em_dashes": len(re.findall(r"—", content)),
        "en_dashes": len(re.findall(r"–", content)),
        "exclamation_marks": len(re.findall(r"!", content)),
        "full_stops": len(re.findall(r"\.", content)) - content.count("..."),
        "hyphens": len(re.findall(r"-", content)),
        "other_punctuation_marks": len(re.findall(r"[*&%$@]", content)),
        "question_marks": len(re.findall(r"\?", content)),
        "round_brackets": len(re.findall(r"\(|\)", content)),
        "semicolons": len(re.findall(r";", content)),
        "slashes": len(re.findall(r"/", content)),
        "square_brackets": len(re.findall(r"\[|\]", content)),
        "vertical_bars": len(re.findall(r"\|", content)),
    }

# Count total words
def count_words(content):
    return len(re.findall(r"\b\w+\b", content))

# Plot line graph with different lines for each selected punctuation type
def plot_line_graph(df, selected_keys):
    fig, ax = plt.subplots(figsize=(12, 6))

    for key in selected_keys:
        ax.plot(df["filename"], df[key], marker='o', label=key)

    ax.set_title("Punctuation Count Comparison", fontsize=16)
    ax.set_xlabel("Document", fontsize=12)
    ax.set_ylabel("Count", fontsize=12)
    ax.legend(fontsize=10)

    # Improve x-axis label formatting
    ax.set_xticks(range(len(df["filename"])))
    ax.set_xticklabels(
        [label.replace(".docx", "").replace("_", " ") for label in df["filename"]],
        rotation=25,
        ha='right',
        fontsize=9,
        wrap=True
    )

    plt.tight_layout()
    return fig


# Streamlit app UI
st.title("📑 Punctuation Analyzer for DOCX Files")

uploaded_files = st.file_uploader("Upload one or more .docx files", type="docx", accept_multiple_files=True)

if uploaded_files:
    results = []

    for file in uploaded_files:
        with st.spinner(f"Analyzing {file.name}..."):
            text = extract_text(file)
            punctuation = count_punctuation(text)
            word_count = count_words(text)

            result = {
                "filename": file.name,
                "word_count": word_count,
                **punctuation
            }
            results.append(result)

    df = pd.DataFrame(results)
    st.success("✅ Analysis Complete!")

    st.subheader("🔍 Punctuation Summary")
    st.dataframe(df)

    # CSV download
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "📥 Download CSV",
        data=csv,
        file_name="punctuation_summary.csv",
        mime="text/csv"
    )

    # Line graph selection
    st.subheader("📈 Visualize Punctuation Across Documents")
    selected = st.multiselect(
        "Select punctuation types to plot:",
        options=PUNCTUATION_KEYS,
        default=["commas", "full_stops", "question_marks"]
    )

    if selected:
        fig = plot_line_graph(df, selected)
        st.pyplot(fig)

        # Download graph as PNG
        buf = BytesIO()
        fig.savefig(buf, format="png")
        st.download_button(
            "📥 Download Graph as PNG",
            data=buf.getvalue(),
            file_name="punctuation_graph.png",
            mime="image/png"
        )
