"""
Unit tests for blade_section_generator.py
"""

import numpy as np
import pytest
import csv
import os
import tempfile
from pathlib import Path

from blade_section_generator import (
    load_airfoil,
    compute_chord,
    compute_twist,
    compute_sweep_x,
    find_le_index,
    compute_normals,
    apply_te_thickening,
    compute_cap_offset,
    transform_section,
    write_station_csv,
    write_sldcrv,
    write_combined_csv,
)


# ─── Fixtures ────────────────────────────────────────────────

@pytest.fixture
def sample_config():
    """Minimal config matching the structure of blade_config.yaml."""
    return {
        'global_geometry': {
            'span': 39.37,
            'root_chord': 4.724,
            'tip_chord': 3.150,
            'ref_axis_xc': 0.25,
            'taper_start_z': 0.0,
            'sweep_start_z': 0.0,
            'sweep_angle_deg': 0.0,
            'twist_start_z': 0.0,
            'twist_rate_deg_per_bs': -10.0,
        }
    }


@pytest.fixture
def simple_airfoil():
    """A simple symmetric diamond-ish airfoil for testing.
    TE upper -> LE -> TE lower, 5 points."""
    return np.array([
        [1.0,  0.01],   # TE upper
        [0.5,  0.06],   # mid upper
        [0.0,  0.0],    # LE
        [0.5, -0.06],   # mid lower
        [1.0, -0.01],   # TE lower
    ])


@pytest.fixture
def cap_config():
    return {
        'cap_te_xc': 0.225,
        'offset_at_le': 0.05,
        'offset_at_cap_te': 0.011,
        'exponent': 3.333,
        'normal_offset_sign': -1,
    }


@pytest.fixture
def airfoil_csv(tmp_path):
    """Write a temporary airfoil CSV and return its path."""
    path = tmp_path / "test_airfoil.csv"
    with open(path, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['x/c', 'y/c'])
        writer.writerow([1.0, 0.01])
        writer.writerow([0.5, 0.06])
        writer.writerow([0.0, 0.0])
        writer.writerow([0.5, -0.06])
        writer.writerow([1.0, -0.01])
    return str(path)


# ─── Airfoil I/O ────────────────────────────────────────────

class TestLoadAirfoil:
    def test_loads_correct_shape(self, airfoil_csv):
        pts = load_airfoil(airfoil_csv)
        assert pts.shape == (5, 2)

    def test_loads_correct_values(self, airfoil_csv):
        pts = load_airfoil(airfoil_csv)
        np.testing.assert_allclose(pts[0], [1.0, 0.01])
        np.testing.assert_allclose(pts[2], [0.0, 0.0])

    def test_skips_header(self, airfoil_csv):
        pts = load_airfoil(airfoil_csv)
        # First data row should be [1.0, 0.01], not the header
        assert pts[0, 0] == 1.0


# ─── Geometry: Chord ────────────────────────────────────────

class TestComputeChord:
    def test_root_chord(self, sample_config):
        chord = compute_chord(0.0, sample_config)
        assert chord == pytest.approx(4.724)

    def test_tip_chord(self, sample_config):
        chord = compute_chord(1.0, sample_config)
        assert chord == pytest.approx(3.150)

    def test_midspan_linear(self, sample_config):
        chord = compute_chord(0.5, sample_config)
        expected = 4.724 + 0.5 * (3.150 - 4.724)
        assert chord == pytest.approx(expected)

    def test_before_taper_start(self, sample_config):
        """When taper_start_z > 0, stations before it get root chord."""
        sample_config['global_geometry']['taper_start_z'] = 20.0
        chord = compute_chord(0.25, sample_config)  # z = 9.84, < 20
        assert chord == pytest.approx(4.724)

    def test_after_taper_start(self, sample_config):
        """When taper_start_z > 0, taper is computed from that point."""
        sample_config['global_geometry']['taper_start_z'] = 20.0
        # At tip (rR=1.0), z=39.37, taper from 20 to 39.37
        chord = compute_chord(1.0, sample_config)
        assert chord == pytest.approx(3.150)


# ─── Geometry: Twist ────────────────────────────────────────

class TestComputeTwist:
    def test_root_no_twist(self, sample_config):
        twist = compute_twist(0.0, sample_config)
        assert twist == 0.0

    def test_tip_twist(self, sample_config):
        twist = compute_twist(1.0, sample_config)
        # rate = -10 deg/bs, z=39.37, twist_start=0
        # twist = -10 * (39.37 - 0) / 39.37 = -10
        assert twist == pytest.approx(-10.0)

    def test_midspan_twist(self, sample_config):
        twist = compute_twist(0.5, sample_config)
        assert twist == pytest.approx(-5.0)

    def test_before_twist_start(self, sample_config):
        sample_config['global_geometry']['twist_start_z'] = 30.0
        twist = compute_twist(0.25, sample_config)  # z = 9.84 < 30
        assert twist == 0.0


# ─── Geometry: Sweep ────────────────────────────────────────

class TestComputeSweep:
    def test_zero_sweep_angle(self, sample_config):
        sweep = compute_sweep_x(1.0, sample_config)
        assert sweep == pytest.approx(0.0)

    def test_nonzero_sweep(self, sample_config):
        sample_config['global_geometry']['sweep_angle_deg'] = 5.0
        sweep = compute_sweep_x(1.0, sample_config)
        expected = np.tan(np.radians(5.0)) * 39.37
        assert sweep == pytest.approx(expected)

    def test_before_sweep_start(self, sample_config):
        sample_config['global_geometry']['sweep_angle_deg'] = 5.0
        sample_config['global_geometry']['sweep_start_z'] = 30.0
        sweep = compute_sweep_x(0.25, sample_config)  # z=9.84 < 30
        assert sweep == 0.0


# ─── Find LE Index ──────────────────────────────────────────

class TestFindLeIndex:
    def test_simple_airfoil(self, simple_airfoil):
        assert find_le_index(simple_airfoil) == 2

    def test_le_at_zero(self):
        pts = np.array([[0.5, 0.0], [0.0, 0.0], [0.5, 0.0]])
        assert find_le_index(pts) == 1


# ─── Normals ────────────────────────────────────────────────

class TestComputeNormals:
    def test_output_shape(self, simple_airfoil):
        normals = compute_normals(simple_airfoil)
        assert normals.shape == simple_airfoil.shape

    def test_unit_length(self, simple_airfoil):
        normals = compute_normals(simple_airfoil)
        lengths = np.linalg.norm(normals, axis=1)
        # All non-degenerate points should have unit normals
        for length in lengths:
            if length > 1e-6:
                assert length == pytest.approx(1.0, abs=1e-6)

    def test_outward_direction(self, simple_airfoil):
        """Normals on the upper surface should point upward (positive y)."""
        normals = compute_normals(simple_airfoil)
        # Point index 1 is mid-upper, normal should have positive y component
        assert normals[1, 1] > 0

    def test_lower_points_outward(self, simple_airfoil):
        """Normals on the lower surface should point downward (negative y)."""
        normals = compute_normals(simple_airfoil)
        # Point index 3 is mid-lower
        assert normals[3, 1] < 0


# ─── TE Thickening ──────────────────────────────────────────

class TestApplyTeThickening:
    def test_zero_requirement_returns_copy(self, simple_airfoil):
        result = apply_te_thickening(simple_airfoil, 4.0, 0.0)
        np.testing.assert_array_equal(result, simple_airfoil)

    def test_negative_requirement_returns_copy(self, simple_airfoil):
        result = apply_te_thickening(simple_airfoil, 4.0, -1.0)
        np.testing.assert_array_equal(result, simple_airfoil)

    def test_already_thick_enough(self, simple_airfoil):
        """If current TE thickness exceeds requirement, airfoil is unchanged."""
        chord = 4.0
        # Current TE thickness = |0.01 - (-0.01)| * 4.0 = 0.08
        result = apply_te_thickening(simple_airfoil, chord, 0.05)
        np.testing.assert_array_equal(result, simple_airfoil)

    def test_returns_copy_not_reference(self, simple_airfoil):
        result = apply_te_thickening(simple_airfoil, 4.0, 0.0)
        result[0, 0] = 999.0
        assert simple_airfoil[0, 0] != 999.0


# ─── Cap Offset ─────────────────────────────────────────────

class TestComputeCapOffset:
    def test_aft_of_cap_te_unchanged(self, simple_airfoil, cap_config):
        """Points aft of cap_te_xc should have zero offset."""
        normals = compute_normals(simple_airfoil)
        chord = 4.0
        cap_pts = compute_cap_offset(simple_airfoil, normals, cap_config, chord)

        # TE points (x/c=1.0) are aft of cap_te_xc=0.225 -> unchanged
        np.testing.assert_array_equal(cap_pts[0], simple_airfoil[0])
        np.testing.assert_array_equal(cap_pts[-1], simple_airfoil[-1])

    def test_le_has_offset(self, simple_airfoil, cap_config):
        """LE point should be offset."""
        normals = compute_normals(simple_airfoil)
        chord = 4.0
        cap_pts = compute_cap_offset(simple_airfoil, normals, cap_config, chord)

        # LE is at index 2 (x/c=0.0), should differ from base
        le_diff = np.linalg.norm(cap_pts[2] - simple_airfoil[2])
        assert le_diff > 0

    def test_offset_shape_preserved(self, simple_airfoil, cap_config):
        normals = compute_normals(simple_airfoil)
        cap_pts = compute_cap_offset(simple_airfoil, normals, cap_config, 4.0)
        assert cap_pts.shape == simple_airfoil.shape


# ─── Transform Section ──────────────────────────────────────

class TestTransformSection:
    def test_output_shape(self, simple_airfoil):
        pts_3d = transform_section(simple_airfoil, 4.0, 0.0, 0.0, 0.25, 10.0)
        assert pts_3d.shape == (5, 3)

    def test_z_position(self, simple_airfoil):
        pts_3d = transform_section(simple_airfoil, 4.0, 0.0, 0.0, 0.25, 10.0)
        np.testing.assert_allclose(pts_3d[:, 2], 10.0)

    def test_no_twist_scaling(self, simple_airfoil):
        """With zero twist and sweep, X should be chord * x/c - ref_offset."""
        chord = 4.0
        ref_xc = 0.25
        pts_3d = transform_section(simple_airfoil, chord, 0.0, 0.0, ref_xc, 0.0)
        expected_x = simple_airfoil[:, 0] * chord - ref_xc * chord
        np.testing.assert_allclose(pts_3d[:, 0], expected_x, atol=1e-10)

    def test_sweep_offset(self, simple_airfoil):
        """Sweep should shift all X positions."""
        sweep = 1.5
        pts_no_sweep = transform_section(simple_airfoil, 4.0, 0.0, 0.0, 0.25, 0.0)
        pts_sweep = transform_section(simple_airfoil, 4.0, 0.0, sweep, 0.25, 0.0)
        np.testing.assert_allclose(pts_sweep[:, 0] - pts_no_sweep[:, 0], sweep, atol=1e-10)

    def test_twist_rotates(self, simple_airfoil):
        """With twist, points should rotate around the ref axis."""
        pts_no_twist = transform_section(simple_airfoil, 4.0, 0.0, 0.0, 0.25, 0.0)
        pts_twist = transform_section(simple_airfoil, 4.0, 10.0, 0.0, 0.25, 0.0)
        # Points should differ
        assert not np.allclose(pts_no_twist[:, :2], pts_twist[:, :2])

    def test_zero_twist_preserves_y(self, simple_airfoil):
        """With zero twist, Y should be chord * y/c."""
        chord = 4.0
        pts_3d = transform_section(simple_airfoil, chord, 0.0, 0.0, 0.25, 0.0)
        expected_y = simple_airfoil[:, 1] * chord
        np.testing.assert_allclose(pts_3d[:, 1], expected_y, atol=1e-10)


# ─── Output Writers ─────────────────────────────────────────

class TestWriters:
    def test_write_station_csv(self, tmp_path):
        pts = np.array([[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]])
        path = tmp_path / "test.csv"
        write_station_csv(str(path), pts, 1)

        with open(path) as f:
            reader = csv.reader(f)
            header = next(reader)
            assert header == ['X_in', 'Y_in', 'Z_in']
            rows = list(reader)
            assert len(rows) == 2

    def test_write_sldcrv(self, tmp_path):
        pts = np.array([[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]])
        path = tmp_path / "test.txt"
        write_sldcrv(str(path), pts)

        with open(path) as f:
            lines = f.readlines()
            assert len(lines) == 2
            parts = lines[0].strip().split('\t')
            assert len(parts) == 3

    def test_write_combined_csv(self, tmp_path):
        stations = [
            {'id': 1, 'rR': 0.0, 'pts_3d': np.array([[1.0, 2.0, 3.0]])},
            {'id': 2, 'rR': 0.5, 'pts_3d': np.array([[4.0, 5.0, 6.0]])},
        ]
        path = tmp_path / "combined.csv"
        write_combined_csv(str(path), stations)

        with open(path) as f:
            reader = csv.reader(f)
            header = next(reader)
            assert 'Station' in header
            rows = list(reader)
            assert len(rows) == 2


# ─── Integration: Real Airfoil Data ─────────────────────────

class TestWithRealAirfoil:
    """Tests using the actual RC310 airfoil data, if available."""

    @pytest.fixture
    def rc310_path(self):
        path = Path(__file__).parent / "airfoil_RC310.csv"
        if not path.exists():
            pytest.skip("airfoil_RC310.csv not found")
        return str(path)

    def test_load_real_airfoil(self, rc310_path):
        pts = load_airfoil(rc310_path)
        assert pts.shape[1] == 2
        assert len(pts) > 10

    def test_le_is_near_zero(self, rc310_path):
        pts = load_airfoil(rc310_path)
        le_idx = find_le_index(pts)
        assert pts[le_idx, 0] < 0.01  # LE x/c should be near 0

    def test_normals_all_unit_length(self, rc310_path):
        pts = load_airfoil(rc310_path)
        normals = compute_normals(pts)
        lengths = np.linalg.norm(normals, axis=1)
        # Skip degenerate points (coincident neighbors produce zero normals)
        nonzero = lengths > 1e-6
        np.testing.assert_allclose(lengths[nonzero], 1.0, atol=1e-6)

    def test_full_transform_pipeline(self, rc310_path, sample_config):
        """Smoke test: load airfoil, apply TE thickening, transform."""
        pts = load_airfoil(rc310_path)
        chord = compute_chord(0.5, sample_config)
        twist = compute_twist(0.5, sample_config)
        sweep = compute_sweep_x(0.5, sample_config)

        section = apply_te_thickening(pts, chord, 0.05)
        pts_3d = transform_section(section, chord, twist, sweep, 0.25, 19.685)

        assert pts_3d.shape[1] == 3
        assert np.all(np.isfinite(pts_3d))
        # Z should all be the station position
        np.testing.assert_allclose(pts_3d[:, 2], 19.685)
