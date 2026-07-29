"""
analyze_results.py
-------------------
Generates two charts from data/output/matched_employers.xlsx that justify
the matching methodology's key design decision (the similarity threshold):

  1. assets/similarity_score_distribution.png
     Distribution of fuzzy-match scores, split by valid/invalid, so the
     separation between "clearly the same employer" and "clearly not" is
     visible at a glance.

  2. assets/threshold_sensitivity.png
     How many File A employers would be counted as matched at each
     candidate threshold from 50 to 100 - the precision/recall trade-off
     that justifies why 90 was chosen over, say, 80 or 95.

Run after match_employers.py:
    python scripts/analyze_results.py
"""

import pandas as pd
import matplotlib.pyplot as plt

CURRENT_THRESHOLD = 90


def plot_score_distribution(df: pd.DataFrame) -> None:
    valid = df.loc[df["valid_match"], "similarity_score"]
    invalid = df.loc[~df["valid_match"], "similarity_score"]

    fig, ax = plt.subplots(figsize=(8, 5))
    bins = range(30, 105, 5)
    ax.hist(invalid, bins=bins, alpha=0.75, label="Invalid match (< 90)", color="#c0392b")
    ax.hist(valid, bins=bins, alpha=0.75, label="Valid match (\u2265 90)", color="#27ae60")
    ax.axvline(CURRENT_THRESHOLD, color="black", linestyle="--", linewidth=1.2,
                label=f"Threshold = {CURRENT_THRESHOLD}")

    ax.set_title("Distribution of Fuzzy-Match Similarity Scores")
    ax.set_xlabel("Similarity score (token-sort ratio)")
    ax.set_ylabel("Number of File A employers")
    ax.legend()
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig("assets/similarity_score_distribution.png", dpi=150)
    plt.close(fig)


def plot_threshold_sensitivity(df: pd.DataFrame) -> None:
    thresholds = list(range(50, 101, 2))
    match_counts = [(df["similarity_score"] >= t).sum() for t in thresholds]

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.plot(thresholds, match_counts, color="#2c3e50", linewidth=2, marker="o", markersize=3)
    ax.axvline(CURRENT_THRESHOLD, color="#27ae60", linestyle="--", linewidth=1.2,
                label=f"Chosen threshold = {CURRENT_THRESHOLD}")

    ax.set_title("Match Count by Similarity Threshold")
    ax.set_xlabel("Similarity threshold")
    ax.set_ylabel("File A employers counted as matched")
    ax.legend()
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig("assets/threshold_sensitivity.png", dpi=150)
    plt.close(fig)


if __name__ == "__main__":
    df = pd.read_excel("data/output/matched_employers.xlsx")
    plot_score_distribution(df)
    plot_threshold_sensitivity(df)
    print("Saved assets/similarity_score_distribution.png")
    print("Saved assets/threshold_sensitivity.png")
