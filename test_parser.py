from openppt_action import build_pptx, parse_outline

SAMPLE = """Here's your deck:

# Q3 Results
## Revenue up 12% YoY

# Key Wins
- Launched EU region
- Churn down to 2.1%
  - Best quarter ever

# Next Quarter
1. Ship mobile app
* Hire 2 SREs
"""


def test_parse_outline():
    slides = parse_outline(SAMPLE)
    assert len(slides) == 3
    assert slides[0] == {"title": "Q3 Results", "subtitle": "Revenue up 12% YoY", "bullets": []}
    assert slides[1]["bullets"] == [
        (0, "Launched EU region"),
        (0, "Churn down to 2.1%"),
        (1, "Best quarter ever"),
    ]
    assert slides[2]["bullets"] == [(0, "Ship mobile app"), (0, "Hire 2 SREs")]
    assert parse_outline("no headings here") == []


def test_build_pptx():
    data = build_pptx(parse_outline(SAMPLE))
    assert data[:2] == b"PK"  # .pptx is a zip


if __name__ == "__main__":
    test_parse_outline()
    test_build_pptx()
    print("ok")
