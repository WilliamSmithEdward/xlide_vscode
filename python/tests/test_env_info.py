from xlide.env_info import get_package_versions


def test_reports_installed_tracked_packages():
    result = get_package_versions()
    packages = result["packages"]
    # Both runtime dependencies are installed in the test venv.
    assert "pyOpenVBA" in packages
    assert "openpyxl" in packages
    for version in packages.values():
        assert isinstance(version, str) and version


def test_never_raises_for_missing_packages(monkeypatch):
    from xlide import env_info

    def boom(_name):
        raise env_info.metadata.PackageNotFoundError

    monkeypatch.setattr(env_info.metadata, "version", boom)
    assert env_info.get_package_versions() == {"packages": {}}
