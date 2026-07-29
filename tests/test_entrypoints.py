from boq_studio import __version__
from boq_studio.app import main


def test_package_version_matches_release() -> None:
    assert __version__ == "0.1.5"


def test_gui_entrypoint_is_callable() -> None:
    assert callable(main)
