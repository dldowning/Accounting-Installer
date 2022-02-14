from menu_funcs import generate
import sys
import os

main_path = r"data"

new_summary_template_path = os.path.join(main_path, r"new_summary_template.xlsx")
reviewers_aid_template = os.path.join(main_path, r"reviewer_aid_template.xlsx")
summary_template = os.path.join(main_path, r"summary_template.xlsx")
dictionary_template_path = os.path.join(main_path, r"Dictionary_template.xlsx")


def generate_summary():
    generate(
        new_summary_template_path,
        reviewers_aid_template,
        summary_template,
        dictionary_template_path,
        func="summary",
    )


def generate_review():
    generate(
        new_summary_template_path,
        reviewers_aid_template,
        summary_template,
        dictionary_template_path,
        func="review",
    )


def generate_new_summary():
    generate(
        new_summary_template_path,
        reviewers_aid_template,
        summary_template,
        dictionary_template_path,
        func="new_summary",
    )


def generate_map_dict():
    generate(
        new_summary_template_path,
        reviewers_aid_template,
        summary_template,
        dictionary_template_path,
        func="map_dict",
    )


if __name__ == "__main__":

    if len(sys.argv) > 1:
        if sys.argv[1] == "summary":
            generate_summary()
        if sys.argv[1] == "review":
            generate_review()
        if sys.argv[1] == "new_summary":
            generate_new_summary()
        if sys.argv[1] == "map_dict":
            generate_map_dict()

