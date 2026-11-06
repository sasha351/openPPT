import asyncio
import io

from pptx import Presentation

from openppt_filter import (
    SENTINEL,
    Filter,
    _describe_layout,
    _layouts_from_bytes,
    _looks_like_deck_request,
)


def _inlet(f, body, **kw):
    return asyncio.run(f.inlet(body, **kw))


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
        f, "_find_template_bytes", lambda body, chat_id=None: _default_template_bytes()
    )
    body = {
        "messages": [
            {"role": "system", "content": "You are helpful."},
            {"role": "user", "content": "Here is my content... make a deck."},
        ]
    }
    out = _inlet(f, body)
    roles = [m["role"] for m in out["messages"]]
    # merged into the existing system prompt, never a second system message:
    # Gemma's and Nemotron's chat templates only take one, at position 0
    assert roles == ["system", "user"]
    primer = out["messages"][0]["content"]
    assert primer.startswith("You are helpful.")
    assert SENTINEL in primer
    assert 'layout="' in primer
    assert "Section Header" in primer  # a real template layout name
    assert "concrete" in primer  # content-quality bar


def test_inlet_noop_without_template(monkeypatch):
    f = Filter()
    monkeypatch.setattr(f, "_find_template_bytes", lambda body, chat_id=None: None)
    body = {"messages": [{"role": "user", "content": "hi"}]}
    out = _inlet(f, body)
    assert len(out["messages"]) == 1


def test_looks_like_deck_request():
    assert _looks_like_deck_request(
        [{"role": "user", "content": "make me a deck about Q3"}]
    )
    assert _looks_like_deck_request(
        [{"role": "user", "content": "I need a presentation on onboarding"}]
    )
    assert _looks_like_deck_request(
        [{"role": "user", "content": "turn this into slides please"}]
    )
    assert not _looks_like_deck_request([{"role": "user", "content": "hi"}])
    assert not _looks_like_deck_request(
        [{"role": "user", "content": "what's the capital of France?"}]
    )
    assert not _looks_like_deck_request([])


def test_inlet_primes_without_template_on_deck_request(monkeypatch):
    f = Filter()
    monkeypatch.setattr(f, "_find_template_bytes", lambda body, chat_id=None: None)
    body = {"messages": [{"role": "user", "content": "make me a slide deck about Q3"}]}
    out = _inlet(f, body)
    roles = [m["role"] for m in out["messages"]]
    assert roles == ["system", "user"]
    primer = out["messages"][0]["content"]
    assert SENTINEL in primer
    assert 'layout="' not in primer
    assert "concrete" in primer


def test_inlet_respects_prime_without_template_valve_disabled(monkeypatch):
    f = Filter()
    f.valves.prime_without_template = False
    monkeypatch.setattr(f, "_find_template_bytes", lambda body, chat_id=None: None)
    body = {"messages": [{"role": "user", "content": "make me a slide deck about Q3"}]}
    out = _inlet(f, body)
    assert len(out["messages"]) == 1


def test_inlet_does_not_double_prime(monkeypatch):
    f = Filter()
    monkeypatch.setattr(
        f, "_find_template_bytes", lambda body, chat_id=None: _default_template_bytes()
    )
    body = {"messages": [{"role": "user", "content": "make a deck"}]}
    once = _inlet(f, body)
    twice = _inlet(f, once)
    primers = [
        m for m in twice["messages"] if SENTINEL in (m.get("content") or "")
    ]
    assert len(primers) == 1


def test_inlet_never_raises_on_bad_template(monkeypatch):
    f = Filter()
    monkeypatch.setattr(f, "_find_template_bytes", lambda body, chat_id=None: b"not a pptx")
    body = {"messages": [{"role": "user", "content": "make a deck"}]}
    out = _inlet(f, body)  # must swallow the error and pass through
    assert out["messages"][-1]["content"] == "make a deck"


def test_describe_layout_points_two_column_layouts_at_columns():
    prs = Presentation()
    by_name = {lay.name: lay for lay in prs.slide_layouts}
    assert "<column>" in _describe_layout(by_name["Two Content"])


def test_primer_teaches_columns_and_final_answer_rule(monkeypatch):
    f = Filter()
    monkeypatch.setattr(f, "_find_template_bytes", lambda body, chat_id=None: None)
    out = _inlet(f, {"messages": [{"role": "user", "content": "make me a deck"}]})
    primer = out["messages"][0]["content"]
    assert "<column" in primer and "final answer" in primer


def test_follow_up_edit_to_a_deck_is_primed(monkeypatch):
    """'make slide 3 shorter' names no deck keyword; without the primer the
    model drifts back to prose or Markdown and the export finds nothing."""
    f = Filter()
    monkeypatch.setattr(f, "_find_template_bytes", lambda body, chat_id=None: None)
    body = {"messages": [
        {"role": "user", "content": "deck about Q3"},
        {"role": "assistant", "content": '<slide title="Q3"></slide>'},
        {"role": "user", "content": "make the second one shorter"},
    ]}
    out = _inlet(f, body)
    assert out["messages"][0]["role"] == "system"
    assert SENTINEL in out["messages"][0]["content"]
    # an ordinary chat with no outline in it stays untouched
    plain = {"messages": [{"role": "user", "content": "hi"}, {"role": "assistant", "content": "hello"},
                          {"role": "user", "content": "shorter please"}]}
    assert len(_inlet(f, plain)["messages"]) == 3


def test_multimodal_user_content_is_understood(monkeypatch):
    f = Filter()
    monkeypatch.setattr(f, "_find_template_bytes", lambda body, chat_id=None: None)
    body = {"messages": [{"role": "user", "content": [
        {"type": "text", "text": "turn this chart into slides"},
        {"type": "image_url", "image_url": {"url": "data:image/png;base64,AA=="}},
    ]}]}
    out = _inlet(f, body)
    assert SENTINEL in out["messages"][0]["content"]


def test_list_shaped_system_prompt_gets_a_text_part(monkeypatch):
    f = Filter()
    monkeypatch.setattr(f, "_find_template_bytes", lambda body, chat_id=None: None)
    body = {"messages": [
        {"role": "system", "content": [{"type": "text", "text": "Be brief."}]},
        {"role": "user", "content": "slides please"},
    ]}
    out = _inlet(f, body)
    parts = out["messages"][0]["content"]
    assert parts[0]["text"] == "Be brief." and SENTINEL in parts[-1]["text"]
    assert len(out["messages"]) == 2


def _fake_open_webui(async_api, template_path, saved_files=None):
    """Stand-in open_webui modules: Files/Chats async (as since the async-DB
    move, v0.12 included) or sync (older)."""
    import sys
    import types

    class Record:
        filename = "brand.potx"
        path = template_path

    def maybe_async(value):
        async def _get():
            return value
        return _get() if async_api else value

    class Files:
        @staticmethod
        def get_file_by_id(fid):
            return maybe_async(Record() if fid == "t1" else None)

    class Storage:
        @staticmethod
        def get_file(path):
            return path

    class Chat:
        chat = {"messages": [{"role": "user", "content": "deck", "files": saved_files or []}]}

    class Chats:
        @staticmethod
        def get_chat_by_id(cid):
            return maybe_async(Chat() if cid == "c1" else None)

    mods = {
        "open_webui": types.ModuleType("open_webui"),
        "open_webui.models": types.ModuleType("open_webui.models"),
        "open_webui.models.files": types.ModuleType("open_webui.models.files"),
        "open_webui.models.chats": types.ModuleType("open_webui.models.chats"),
        "open_webui.storage": types.ModuleType("open_webui.storage"),
        "open_webui.storage.provider": types.ModuleType("open_webui.storage.provider"),
    }
    mods["open_webui.models.files"].Files = Files
    mods["open_webui.models.chats"].Chats = Chats
    mods["open_webui.storage.provider"].Storage = Storage
    sys.modules.update(mods)
    return lambda: [sys.modules.pop(k, None) for k in mods]


def test_template_is_found_with_the_async_files_api():
    """v0.12's Files.get_file_by_id is async; the sync filter never awaited
    it, so the template was never found and the layouts never primed."""
    import tempfile

    with tempfile.NamedTemporaryFile(suffix=".potx", delete=False) as fh:
        fh.write(_default_template_bytes())
    for async_api in (True, False):
        undo = _fake_open_webui(async_api, fh.name)
        try:
            body = {"files": [{"type": "file", "id": "t1", "name": "brand.potx"}],
                    "messages": [{"role": "user", "content": "turn my notes into a deck"}]}
            out = _inlet(Filter(), body)
        finally:
            undo()
        primer = out["messages"][0]["content"]
        assert 'layout="' in primer and "Section Header" in primer, async_api


def test_template_attached_turns_ago_is_found_in_the_saved_chat():
    import tempfile

    with tempfile.NamedTemporaryFile(suffix=".potx", delete=False) as fh:
        fh.write(_default_template_bytes())
    undo = _fake_open_webui(True, fh.name, saved_files=[{"id": "t1", "name": "brand.potx"}])
    try:
        body = {"messages": [{"role": "user", "content": "now add a risks slide"}]}
        out = _inlet(Filter(), body, __chat_id__="c1")
    finally:
        undo()
    assert "Section Header" in out["messages"][0]["content"]


if __name__ == "__main__":
    import inspect

    class _MP:
        """Just enough of pytest's monkeypatch fixture to run standalone."""

        def setattr(self, obj, name, val):
            setattr(obj, name, val)

    # Auto-discover like test_parser.py, so a new test can't be forgotten here.
    for _name, _fn in sorted(list(globals().items())):
        if _name.startswith("test_") and callable(_fn):
            params = inspect.signature(_fn).parameters
            _fn(*([_MP()] if "monkeypatch" in params else []))
    print("ok")
