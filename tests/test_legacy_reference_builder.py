import pytest

from tools import build_reference


def test_legacy_reference_builder_is_explicitly_disabled():
    with pytest.raises(RuntimeError, match="已停用"):
        build_reference.main()
