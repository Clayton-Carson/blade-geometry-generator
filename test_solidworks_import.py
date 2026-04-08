"""
Unit tests for solidworks_import.py (data-loading and helper functions only).
SolidWorks COM automation is not tested here.
"""

import numpy as np
import pytest
from pathlib import Path

from solidworks_import import (
    read_curve_points,
    find_le_index,
    get_lower_surface,
    interpolate_lower_at_xc,
    discover_curve_files,
    build_guide_points,
    INCHES_TO_METERS,
)


# ─── Fixtures ────────────────────────────────────────────────

@pytest.fixture
def curve_file(tmp_path):
    """Write a tab-separated curve file and return its path."""
    path = tmp_path / "test_curve.txt"
    with open(path, 'w') as f:
        f.write("3.00000000\t0.10000000\t5.00000000\n")
        f.write("1.50000000\t0.30000000\t5.00000000\n")
        f.write("-0.50000000\t0.00000000\t5.00000000\n")
        f.write("1.50000000\t-0.30000000\t5.00000000\n")
        f.write("3.00000000\t-0.10000000\t5.00000000\n")
    return path


@pytest.fixture
def curves_dir(tmp_path):
    """Create a directory with 8 base and 8 cap curve files.
    Each file has 30 points (enough for UPPER_GUIDE_INDICES up to 24).
    """
    d = tmp_path / "curves"
    d.mkdir()
    for prefix in ['base', 'cap']:
        for i in range(1, 9):
            path = d / f"{prefix}_S{i:02d}.txt"
            z = float(i) * 5.0
            with open(path, 'w') as f:
                # 30-point airfoil: TE upper -> LE -> TE lower
                n_half = 15
                for k in range(n_half):
                    # Upper surface: TE to LE (x decreasing)
                    t = k / (n_half - 1)
                    x = 3.0 * (1.0 - t) - 0.5 * t
                    y = 0.3 * np.sin(np.pi * t)
                    f.write(f"{x:.8f}\t{y:.8f}\t{z:.8f}\n")
                for k in range(1, n_half):
                    # Lower surface: LE to TE (x increasing)
                    t = k / (n_half - 1)
                    x = -0.5 * (1.0 - t) + 3.0 * t
                    y = -0.3 * np.sin(np.pi * t)
                    f.write(f"{x:.8f}\t{y:.8f}\t{z:.8f}\n")
    return d


# ─── Data Loading ───────────────────────────────────────────

class TestReadCurvePoints:
    def test_shape(self, curve_file):
        pts = read_curve_points(curve_file)
        assert pts.shape == (5, 3)

    def test_values(self, curve_file):
        pts = read_curve_points(curve_file)
        np.testing.assert_allclose(pts[0], [3.0, 0.1, 5.0])
        np.testing.assert_allclose(pts[2], [-0.5, 0.0, 5.0])

    def test_empty_file(self, tmp_path):
        path = tmp_path / "empty.txt"
        path.write_text("")
        pts = read_curve_points(path)
        assert len(pts) == 0


class TestFindLeIndex:
    def test_simple(self):
        pts = np.array([[3.0, 0.1, 0.0], [0.0, 0.0, 0.0], [3.0, -0.1, 0.0]])
        assert find_le_index(pts) == 1

    def test_negative_x(self):
        pts = np.array([[3.0, 0.1, 0.0], [-0.5, 0.0, 0.0], [3.0, -0.1, 0.0]])
        assert find_le_index(pts) == 1


class TestGetLowerSurface:
    def test_returns_from_le_to_end(self):
        pts = np.array([
            [3.0, 0.1, 0.0],
            [0.0, 0.0, 0.0],
            [1.5, -0.3, 0.0],
            [3.0, -0.1, 0.0],
        ])
        lower = get_lower_surface(pts)
        assert len(lower) == 3  # LE + 2 lower points
        np.testing.assert_array_equal(lower[0], pts[1])  # starts at LE


class TestInterpolateLowerAtXc:
    def test_at_le(self):
        lower = np.array([
            [0.0, 0.0, 5.0],
            [1.5, -0.3, 5.0],
            [3.0, -0.1, 5.0],
        ])
        pt = interpolate_lower_at_xc(lower, 0.0)
        np.testing.assert_allclose(pt, [0.0, 0.0, 5.0])

    def test_at_te(self):
        lower = np.array([
            [0.0, 0.0, 5.0],
            [1.5, -0.3, 5.0],
            [3.0, -0.1, 5.0],
        ])
        pt = interpolate_lower_at_xc(lower, 1.0)
        np.testing.assert_allclose(pt, [3.0, -0.1, 5.0])

    def test_midpoint(self):
        lower = np.array([
            [0.0, 0.0, 5.0],
            [3.0, -0.6, 5.0],
        ])
        pt = interpolate_lower_at_xc(lower, 0.5)
        np.testing.assert_allclose(pt, [1.5, -0.3, 5.0])


# ─── File Discovery ─────────────────────────────────────────

class TestDiscoverCurveFiles:
    def test_finds_all_files(self, curves_dir):
        files = discover_curve_files(curves_dir, 'base')
        assert len(files) == 8
        assert all(f.exists() for f in files)

    def test_correct_order(self, curves_dir):
        files = discover_curve_files(curves_dir, 'base')
        assert files[0].name == 'base_S01.txt'
        assert files[7].name == 'base_S08.txt'

    def test_missing_file_raises(self, tmp_path):
        d = tmp_path / "incomplete"
        d.mkdir()
        # Only create 3 of 8 files
        for i in range(1, 4):
            (d / f"base_S{i:02d}.txt").write_text("0\t0\t0\n")
        with pytest.raises(FileNotFoundError):
            discover_curve_files(d, 'base')


# ─── Guide Points ───────────────────────────────────────────

class TestBuildGuidePoints:
    def test_returns_expected_keys(self, curves_dir):
        guides = build_guide_points(curves_dir, 'base')
        assert 'TE_upper' in guides
        assert 'TE_lower' in guides
        assert 'LE' in guides

    def test_te_upper_is_first_point(self, curves_dir):
        guides = build_guide_points(curves_dir, 'base')
        # TE upper should be first point of each station
        assert guides['TE_upper'].shape == (8, 3)
        # First station Z = 5.0
        assert guides['TE_upper'][0, 2] == pytest.approx(5.0)

    def test_te_lower_is_last_point(self, curves_dir):
        guides = build_guide_points(curves_dir, 'base')
        assert guides['TE_lower'].shape == (8, 3)

    def test_le_is_min_x(self, curves_dir):
        guides = build_guide_points(curves_dir, 'base')
        # LE x should be -0.5 for all stations (min X in our test data)
        np.testing.assert_allclose(guides['LE'][:, 0], -0.5)

    def test_z_increases_across_stations(self, curves_dir):
        guides = build_guide_points(curves_dir, 'base')
        z_vals = guides['LE'][:, 2]
        assert all(z_vals[i] < z_vals[i + 1] for i in range(len(z_vals) - 1))


# ─── Constants ──────────────────────────────────────────────

class TestConstants:
    def test_inches_to_meters(self):
        assert INCHES_TO_METERS == pytest.approx(0.0254)
