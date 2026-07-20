from openppt_action import build_pptx, parse_outline

SAMPLE = """Here's your deck:

# Q3 Results
## Revenue up 12% YoY

# Key Wins
- Launched EU region
- Churn down to 2.1%
  - Best quarter ever
Notes: mention the EU launch timeline if asked

# Next Quarter
1. Ship mobile app
* Hire 2 SREs

# Team Photo
- ![Team offsite](https://example.com/team.jpg)
"""


def test_parse_outline():
    slides = parse_outline(SAMPLE)
    assert len(slides) == 4
    assert slides[0] == {
        "title": "Q3 Results",
        "subtitle": "Revenue up 12% YoY",
        "notes": "",
        "bullets": [],
    }
    assert slides[1]["bullets"] == [
        (0, "Launched EU region"),
        (0, "Churn down to 2.1%"),
        (1, "Best quarter ever"),
    ]
    assert slides[1]["notes"] == "mention the EU launch timeline if asked"
    assert slides[2]["bullets"] == [(0, "Ship mobile app"), (0, "Hire 2 SREs")]
    assert slides[3]["bullets"] == [
        (0, {"alt": "Team offsite", "image": "https://example.com/team.jpg"})
    ]
    assert parse_outline("no headings here") == []


def test_build_pptx():
    data = build_pptx(parse_outline(SAMPLE))
    assert data[:2] == b"PK"  # .pptx is a zip


def test_build_pptx_skips_unreachable_image():
    # image slide with an unfetchable URL should still produce a valid deck
    slides = parse_outline("# Photo\n- ![x](https://nonexistent.invalid/x.png)\n")
    data = build_pptx(slides)
    assert data[:2] == b"PK"


if __name__ == "__main__":
    test_parse_outline()
    test_build_pptx()
    test_build_pptx_skips_unreachable_image()
    print("ok")
