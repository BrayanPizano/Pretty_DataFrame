import pandas as pd
import numpy as np
import random
import math


def generate_color():
    return tuple(random.randint(50, 200) for _ in range(3))


def get_contrasting_color(rgb):
    r, g, b = rgb
    brightness = (r * 299 + g * 587 + b * 114) / 1000
    return (0, 0, 0) if brightness > 125 else (255, 255, 255)


def ansi_style(text, bg=None, fg=None):
    seq = ""
    if bg:
        seq += f"\033[48;2;{bg[0]};{bg[1]};{bg[2]}m"
    if fg:
        seq += f"\033[38;2;{fg[0]};{fg[1]};{fg[2]}m"
    return f"{seq}{text}\033[0m"


def truncate_text(text, max_length):
    if len(text) <= max_length:
        return text
    return text[:max_length - 3] + "..."


def handle_nan_value(value, nan_placeholder="N/A"):
    """Handle NaN or None values with a placeholder"""
    if pd.isna(value):
        return nan_placeholder
    return str(value)


def draw_table_border(widths, style="ascii", is_header=False):
    """Draw table borders using ASCII or Unicode characters"""
    if style == "none":
        return ""

    # Character sets for different border styles
    border_chars = {
        "ascii": {
            "top_left": "+", "top_right": "+", "bottom_left": "+", "bottom_right": "+",
            "top_middle": "+", "bottom_middle": "+", "left_middle": "+", "right_middle": "+",
            "middle": "+", "horizontal": "-", "vertical": "|"
        },
        "unicode": {
            "top_left": "┌", "top_right": "┐", "bottom_left": "└", "bottom_right": "┘",
            "top_middle": "┬", "bottom_middle": "┴", "left_middle": "├", "right_middle": "┤",
            "middle": "┼", "horizontal": "─", "vertical": "│"
        },
        "unicode_heavy": {
            "top_left": "┏", "top_right": "┓", "bottom_left": "┗", "bottom_right": "┛",
            "top_middle": "┳", "bottom_middle": "┻", "left_middle": "┣", "right_middle": "┫",
            "middle": "╋", "horizontal": "━", "vertical": "┃"
        }
    }

    chars = border_chars.get(style, border_chars["ascii"])

    # Create the appropriate border line based on position
    if is_header:
        middle_char = chars["middle"]
        left_char = chars["left_middle"]
        right_char = chars["right_middle"]
    else:
        if is_header is None:  # Top border
            middle_char = chars["top_middle"]
            left_char = chars["top_left"]
            right_char = chars["top_right"]
        else:  # Bottom border
            middle_char = chars["bottom_middle"]
            left_char = chars["bottom_left"]
            right_char = chars["bottom_right"]

    result = left_char
    for i, width in enumerate(widths):
        result += chars["horizontal"] * (width + 2)  # +2 for padding spaces
        result += middle_char if i < len(widths) - 1 else right_char

    return result


def find_column_with_fewest_unique_values(df):
    """Find the column with the fewest unique values to use as default grouping column"""
    if df.empty or len(df.columns) == 0:
        return None

    # Calculate the number of unique values in each column
    unique_counts = {col: df[col].nunique() for col in df.columns}

    # Find the column with the minimum number of unique values
    min_col = min(unique_counts, key=unique_counts.get)

    # Only use columns with at least 2 unique values and less than half the number of rows
    good_cols = {col: count for col, count in unique_counts.items()
                 if count >= 2 and count < len(df) / 2}

    if good_cols:
        return min(good_cols, key=good_cols.get)

    return min_col


def print_colored_df(
        df,
        group_col=None,
        special_col=None,
        style_mode="background",
        persist_special_color=False,
        max_column_width=20,
        separator="  ",
        border_style="none",
        nan_placeholder="N/A",
        page_size=None,
        current_page=1,
        group_pagination=None
):
    """
    Display a DataFrame with colored formatting and various display options.

    Parameters:
    -----------
    df : pandas.DataFrame
        The DataFrame to display
    group_col : str, optional
        Column name used for grouping rows by color. If None, the column with fewest unique values is used
    special_col : str, optional
        Column name for special highlighting. If None, no special highlighting is applied
    style_mode : str, default="background"
        "background" for colored backgrounds, "text" for colored text
    persist_special_color : bool, default=False
        Whether to apply special color to columns after special_col
    max_column_width : int, default=20
        Maximum width for any column
    separator : str, default="  "
        Separator string between columns
    border_style : str, default="none"
        Border style: "none", "ascii", "unicode", or "unicode_heavy"
    nan_placeholder : str, default="N/A"
        String to display for NaN values
    page_size : int, optional
        Number of rows per page (None for all rows)
    current_page : int, default=1
        Current page to display when using pagination
    group_pagination : str or list, optional
        If provided, pagination will be by group(s) instead of row numbers:
        - If str: Show data for the specified group value
        - If list: Show data for the specified group values
        - If None: Pagination based on page_size and current_page
    """

    # Make a copy of the dataframe to avoid modifying the original
    df_full = df.copy()

    # Auto-select group_col if not specified
    if group_col is None:
        group_col = find_column_with_fewest_unique_values(df_full)
        if group_col:
            print(f"Auto-selected '{group_col}' as group column (fewest unique values: {df_full[group_col].nunique()})")
        else:
            print(f"Warning: Couldn't find a suitable grouping column. Using index instead.")
            df_full.reset_index(inplace=True)
            group_col = 'index'

    # Handle missing group column
    if group_col not in df_full.columns:
        print(f"Warning: group_col '{group_col}' not found in DataFrame. Using index instead.")
        df_full.reset_index(inplace=True)
        group_col = 'index'

    # Handle special_col
    if special_col is not None and special_col not in df_full.columns:
        print(f"Warning: special_col '{special_col}' not found in DataFrame. Not applying special highlighting.")
        special_col = None

    # Process group pagination if specified
    if group_pagination is not None:
        # Get all unique group values for reference
        all_groups = sorted(df_full[group_col].unique())

        if isinstance(group_pagination, (str, int, float)):
            # Single group specified
            group_pagination = [group_pagination]

        # Filter dataframe to only include specified groups
        df_display = df_full[df_full[group_col].isin(group_pagination)].copy()

        if len(df_display) == 0:
            print(f"No data found for group(s): {group_pagination}")
            print(f"Available groups: {all_groups}")
            return

        # Show info about which groups are being displayed
        total_groups = len(all_groups)
        current_groups = group_pagination

        group_info = f"Showing {len(current_groups)} of {total_groups} groups: {current_groups}"

    else:
        # Standard row-based pagination
        total_rows = len(df_full)

        if page_size is not None:
            total_pages = math.ceil(total_rows / page_size)
            current_page = max(1, min(current_page, total_pages))
            start_idx = (current_page - 1) * page_size
            end_idx = min(start_idx + page_size, total_rows)
            df_display = df_full.iloc[start_idx:end_idx].copy()

            # Pagination info for row-based pagination
            group_info = f"Page {current_page} of {total_pages} | Rows {start_idx + 1}-{end_idx} of {total_rows}"
        else:
            df_display = df_full.copy()
            group_info = f"Showing all {len(df_full)} rows"

    # Calculate column widths with maximum limit
    col_widths = {}
    for col in df_display.columns:
        header_len = len(str(col))
        # Handle NaN values when calculating width
        max_data_len = df_display[col].apply(lambda x: len(handle_nan_value(x, nan_placeholder))).max()
        col_widths[col] = min(max(header_len, max_data_len), max_column_width)

    # Widths list for border drawing
    widths_list = [col_widths[col] for col in df_display.columns]

    # Generate colors for groups
    group_colors = {g: generate_color() for g in df_full[group_col].unique()}

    # Generate colors for special values only if special_col is specified
    special_colors = {}
    if special_col is not None:
        special_colors = {s: generate_color() for s in df_full[special_col].dropna().unique()}
        # Add a default color for NaN values in special column
        special_colors[np.nan] = generate_color()

    # Get border characters if needed
    if border_style != "none":
        border_chars = {
            "ascii": {"vertical": "|"},
            "unicode": {"vertical": "│"},
            "unicode_heavy": {"vertical": "┃"}
        }
        border_char = border_chars.get(border_style, {"vertical": "|"})["vertical"]

    # Draw top border if needed
    if border_style != "none":
        print(draw_table_border(widths_list, border_style, is_header=None))

    # Print header
    header_parts = []
    for col in df_display.columns:
        col_str = str(col)
        if len(col_str) > col_widths[col]:
            col_str = truncate_text(col_str, col_widths[col])
        header_parts.append(f"{col_str:<{col_widths[col]}}")

    if border_style != "none":
        header = f"{border_char} " + f" {border_char} ".join(header_parts) + f" {border_char}"
    else:
        header = separator.join(header_parts)

    print(header)

    # Draw header separator if using borders
    if border_style != "none":
        print(draw_table_border(widths_list, border_style, is_header=True))

    # Print rows
    for _, row in df_display.iterrows():
        group_val = row[group_col]

        # Get the base color for this row based on its group
        row_color = group_colors[group_val]
        text_color = get_contrasting_color(row_color)

        row_str = ""

        # Start with left border if needed
        if border_style != "none":
            row_str += border_char + " "

        # Process each column separately
        for i, col in enumerate(df_display.columns):
            # Handle NaN values
            val = handle_nan_value(row[col], nan_placeholder)

            # Truncate value if too long
            if len(val) > col_widths[col]:
                val = truncate_text(val, col_widths[col])

            val = val.ljust(col_widths[col])

            # Set default styling based on group color
            bg = row_color if style_mode == "background" else None
            fg = text_color if style_mode == "background" else row_color

            # Apply special coloring if special_col is specified and this column/cell should be special
            if special_col is not None:
                special_val = row[special_col]
                # Make sure NaN is handled properly for color mapping
                if pd.isna(special_val):
                    special_val = np.nan

                special_color = special_colors[special_val]
                special_text_color = get_contrasting_color(special_color)

                is_special = (col == special_col)
                apply_special = is_special or (persist_special_color and col in df_display.columns[
                                                                                df_display.columns.get_loc(
                                                                                    special_col) + 1:])

                # Set coloring based on mode and column
                if apply_special:
                    if style_mode == "text":
                        fg = special_color
                    else:  # background
                        bg = special_color
                        fg = special_text_color

            # Apply style to the column value
            row_str += ansi_style(val, bg=bg, fg=fg)

            # Add separator or border after the column (except the last one)
            if i < len(df_display.columns) - 1:
                if border_style != "none":
                    # Reset style, then add border character with neutral styling
                    row_str += "\033[0m " + border_char + " "
                else:
                    # Reset style, then add separator with neutral styling
                    row_str += "\033[0m" + separator

        # Add right border if needed
        if border_style != "none":
            row_str += "\033[0m " + border_char

        print(row_str)

    # Draw bottom border if needed
    if border_style != "none":
        print(draw_table_border(widths_list, border_style, is_header=False))

    # Print pagination/group info
    print(f"\n{group_info}")

    # Print pagination/grouping guidance
    if group_pagination is not None:
        all_groups = sorted(df_full[group_col].unique())
        print(f"Available groups: {all_groups}")
        print(f"Use print_colored_df(..., group_pagination=['group1', 'group2']) to view specific groups")
    elif page_size is not None:
        print(f"Use print_colored_df(..., current_page=N) to view other pages")
    print(f"Use print_colored_df(..., group_pagination='group_name') to view by group")


# Example usage with all features including group pagination
def example():
    # Create a test dataframe with some NaN values
    df = pd.DataFrame({
        'Group': ['A'] * 6 + ['B'] * 6 + ['C'] * 6,
        'col1': np.random.randint(0, 100, 18),
        'col2': [np.nan if i % 7 == 0 else np.random.randint(0, 100) for i in range(18)],
        'col3': ["This is a very long text that should be truncated" if i % 5 == 0 else i for i in range(18)],
        'special_col': ['X', 'Y', np.nan, 'V', 'Z', 'V', 'X', 'U', 'W', 'Y', 'Z', 'X', 'V', 'Y', 'Z', 'Z', 'Y', 'X'],
        'col4': np.random.randint(0, 100, 18),
        'col5': np.random.randint(0, 100, 18)
    })

    print("\n1. Basic usage with custom separator:")
    print_colored_df(df, group_col="Group", special_col="special_col",
                     style_mode="text", separator=" | ")

    print("\n2. With ASCII borders:")
    print_colored_df(df, group_col="Group", special_col="special_col",
                     style_mode="background", border_style="ascii")

    print("\n3. With Unicode borders and custom NaN placeholder:")
    print_colored_df(df, group_col="Group", special_col="special_col",
                     style_mode="background", border_style="unicode",
                     nan_placeholder="<NULL>")

    print("\n4. With pagination (page 1):")
    print_colored_df(df, group_col="Group", special_col="special_col",
                     border_style="unicode_heavy", page_size=5, current_page=1)

    print("\n5. With pagination (page 2):")
    print_colored_df(df, group_col="Group", special_col="special_col",
                     border_style="unicode_heavy", page_size=5, current_page=2)

    print("\n6. With group pagination (showing only group 'A'):")
    print_colored_df(df, group_col="Group", special_col="special_col",
                     border_style="unicode", group_pagination="A")

    print("\n7. With multiple group pagination (showing groups 'A' and 'C'):")
    print_colored_df(df, group_col="Group", special_col="special_col",
                     border_style="unicode", group_pagination=["A", "C"])

    print("\n8. Auto-selecting group column (no group_col specified):")
    print_colored_df(df, border_style="unicode")

    print("\n9. No special column specified (everything colored by group):")
    print_colored_df(df, group_col="Group", border_style="unicode")

    # Create a dataframe with more columns that could be sensible group candidates
    df2 = pd.DataFrame({
        'Category': ['Electronics'] * 5 + ['Clothing'] * 3 + ['Books'] * 2,
        'Status': ['In Stock', 'Low Stock', 'Out of Stock'] * 3 + ['In Stock'],
        'Price': np.random.randint(10, 200, 10),
        'Rating': [4.5, 3.8, 4.2, 2.7, 5.0, 4.1, 3.9, 4.8, 3.5, 4.3],
        'Review': ["Good", "Bad", "Excellent", "Poor", "Perfect"] * 2
    })

    print("\n10. Auto-selecting group with multiple candidate columns:")
    print_colored_df(df2, border_style="unicode")


if __name__ == "__main__":
    example()
