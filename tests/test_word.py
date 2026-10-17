import pytest

from jaraco.office import word


def test_handle_multiple(tmp_path, monkeypatch):
    monkeypatch.setattr(word, 'Converter', lambda: bytes.upper)
    for name in 'a', 'b':
        tmp_path.joinpath(f'{name}.doc').write_bytes(name.encode())
    word.handle_multiple(str(tmp_path / '*.doc'))
    assert tmp_path.joinpath('a.pdf').read_bytes() == b'A'
    assert tmp_path.joinpath('b.pdf').read_bytes() == b'B'


def test_handle_single_with_target(tmp_path, monkeypatch):
    monkeypatch.setattr(word, 'Converter', lambda: bytes.upper)
    tmp_path.joinpath('a.doc').write_bytes(b'a')
    target = tmp_path / 'out.pdf'
    word.handle_multiple(str(tmp_path / 'a.doc'), str(target))
    assert target.read_bytes() == b'A'


def test_target_with_multiple_sources(tmp_path):
    for name in 'a', 'b':
        tmp_path.joinpath(f'{name}.doc').write_bytes(b'')
    with pytest.raises(Exception, match='Cannot specify output file'):
        word.handle_multiple(str(tmp_path / '*.doc'), 'out.pdf')


def test_doc_to_pdf_cmd(tmp_path, monkeypatch):
    monkeypatch.setattr(word, 'Converter', lambda: bytes.upper)
    source = tmp_path / 'a.doc'
    source.write_bytes(b'a')
    monkeypatch.setattr('sys.argv', ['doc-to-pdf', str(source)])
    word.doc_to_pdf_cmd()
    assert tmp_path.joinpath('a.pdf').read_bytes() == b'A'
