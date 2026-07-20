import io

from pptx import Presentation

from openppt_filter import SENTINEL, Filter, _describe_layout, _layouts_from_bytes


def _default_template_bytes():
    prs = Presentation()
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def test_describe_layout_hints():
    prs = Presentation()
    by_name = {lay.name: lay for lay in prs.slide_layouts}
    assert "title + subtitle" in _describe_layout(by_name["Title Slide"])
    assert "divider" in _describe_layout(by_name["Title Only"])
    assert "two columns" in _describe_layout(by_name["Two Content"])
    assert "bullet list" in _describe_layout(by_name["Title and Content"])


def test_layouts_from_bytes_lists_template_layouts():
    lines = _layouts_from_bytes(_default_template_bytes())
    joined = "\n".join(lines)
    assert "Section Header" in joined and "Title and Content" in joined
    assert all(line.startswith("- ") for line in lines)


def test_inlet_injects_primer_when_template_present(monkeypatch):
    f = Filter()
    monkeypatch.setattr(
        f, "_find_template_bytes", lambda body: _default_template_bytes()
    )
    body = {
        "messages": [
            {"role": "system", "content": "You are helpful."},
            {"role": "user", "content": "Here is my content... make a deck."},
        ]
    }
    out = f.inlet(body)
    roles = [m["role"] for m in out["messages"]]
    # primer inserted right after the existing system prompt
    assert roles == ["system", "system", "user"]
    primer = out["messages"][1]["content"]
    assert SENTINEL in primer
    assert "@layout:" in primer
    assert "Section Header" in primer  # a real template layout name


def test_inlet_noop_without_template(monkeypatch):
    f = Filter()
    monkeypatch.setattr(f, "_find_template_bytes", lambda body: None)
    body = {"messages": [{"role": "user", "content": "hi"}]}
    out = f.inlet(body)
    assert len(out["messages"]) == 1


def test_inlet_does_not_double_prime(monkeypatch):
    f = Filter()
    monkeypatch.setattr(
        f, "_find_template_bytes", lambda body: _default_template_bytes()
    )
    body = {"messages": [{"role": "user", "content": "make a deck"}]}
    once = f.inlet(body)
    twice = f.inlet(once)
    primers = [
        m for m in twice["messages"] if SENTINEL in (m.get("content") or "")
    ]
    assert len(primers) == 1


def test_inlet_never_raises_on_bad_template(monkeypatch):
    f = Filter()
    monkeypatch.setattr(f, "_find_template_bytes", lambda body: b"not a pptx")
    body = {"messages": [{"role": "user", "content": "make a deck"}]}
    out = f.inlet(body)  # must swallow the error and pass through
    assert out["messages"][-1]["content"] == "make a deck"


if __name__ == "__main__":
    import types

    class _MP:
        def setattr(self, obj, name, val):
            setattr(obj, name, val)

    mp = _MP()
    test_describe_layout_hints()
    test_layouts_from_bytes_lists_template_layouts()
    test_inlet_injects_primer_when_template_present(mp)
    test_inlet_noop_without_template(mp)
    test_inlet_does_not_double_prime(mp)
    test_inlet_never_raises_on_bad_template(mp)
    print("ok")
