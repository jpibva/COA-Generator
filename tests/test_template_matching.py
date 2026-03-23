import unittest

from coa_template_match import score_template_candidate


class TemplateMatchingTests(unittest.TestCase):
    def test_exact_product_and_client_score_higher_than_generic_candidate(self):
        exact = score_template_candidate(
            "Organic Strawberry",
            "Trader Joe's",
            "/templates/Trader Joes/Template_Organic_Strawberry.docx",
            {"organic", "template"},
        )
        generic = score_template_candidate(
            "Organic Strawberry",
            "Trader Joe's",
            "/templates/General/Template_Mix.docx",
            {"organic", "template"},
        )
        self.assertGreater(exact, generic)

    def test_mix_template_is_penalized_for_non_mix_product(self):
        regular = score_template_candidate(
            "Blueberry",
            "Woodland Partners",
            "/templates/Woodland/Template_Blueberry.docx",
            {"template"},
        )
        mix = score_template_candidate(
            "Blueberry",
            "Woodland Partners",
            "/templates/Woodland/Template_Mix.docx",
            {"template"},
        )
        self.assertGreater(regular, mix)


if __name__ == "__main__":
    unittest.main()
