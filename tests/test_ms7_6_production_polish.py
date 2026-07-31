import pytest

from boq_studio.appearance import (
    DARK_PALETTE,
    LIGHT_PALETTE,
    ThemeMode,
    accessible_widget_contract,
    build_profile,
    clamp_scale,
    font_family_for,
    scale_from_dpi,
)
from boq_studio.workspace import WorkspaceServices


def test_platform_typography_uses_native_cross_platform_families():
    assert font_family_for("Windows") == "Segoe UI"
    assert font_family_for("Darwin") == "Helvetica Neue"
    assert font_family_for("Linux") == "DejaVu Sans"


def test_dpi_scaling_is_clamped_for_safe_desktop_layouts():
    assert scale_from_dpi(96) == 1.0
    assert scale_from_dpi(144) == 1.5
    assert scale_from_dpi(384) == 2.0
    assert clamp_scale(0.5) == 1.0


def test_semantic_palettes_keep_light_and_dark_modes_decoupled_from_widgets():
    light = build_profile(system_name="Windows", theme_mode=ThemeMode.LIGHT)
    dark = build_profile(system_name="Windows", theme_mode=ThemeMode.DARK)
    assert light.palette is LIGHT_PALETTE
    assert dark.palette is DARK_PALETTE
    assert light.palette.text != dark.palette.text
    assert light.palette.accent
    assert dark.palette.accent


def test_appearance_profile_exposes_consistent_spacing_and_touch_targets():
    profile = build_profile(system_name="Darwin", dpi=192)
    assert profile.scale == 2.0
    assert profile.spacing.xs < profile.spacing.sm < profile.spacing.md < profile.spacing.lg
    assert profile.minimum_target_px >= 28


def test_accessibility_contract_does_not_depend_on_color_or_mouse_only_input():
    contract = accessible_widget_contract()
    assert contract["keyboard_focus_visible"] is True
    assert contract["text_labels_required"] is True
    assert contract["color_only_state_forbidden"] is True
    assert contract["tooltip_delay_ms"] >= 300


def test_workspace_services_share_one_appearance_profile_across_surfaces():
    profile = build_profile(system_name="Windows")
    services = WorkspaceServices(appearance=profile)
    assert services.appearance is profile


def test_invalid_theme_mode_is_rejected_by_enum():
    with pytest.raises(ValueError):
        ThemeMode("sepia")
