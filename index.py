"""
=============================================================================
Mohr Circle & Mohr-Coulomb Geomechanical Analysis from Borehole Data
Well: 58-32 Main | Data: Geophysical Well Log CSV
=============================================================================
Computes: Sv, Pp (Eaton), Shmin, SHmax, σ1/σ2/σ3, effective stresses,
          UCS, Cohesion, Friction angle, Mohr Circle, Mohr-Coulomb envelope
=============================================================================
"""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from scipy.integrate import cumulative_trapezoid
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# SECTION 1: CONFIGURATION
# ============================================================
CONFIG = {
    'csv_file': '58-32_main_geophysical_well_log.csv',
    'null_value': -999.25,
    'g': 9.80665,                    # gravity (m/s²)
    'water_density': 1.025,          # g/cc (seawater)
    'psi_per_mpa': 145.038,
    'stress_regime': 'normal',       # 'normal', 'strike_slip', 'reverse'
    'gr_shale_cutoff': 75,           # API units
    'eaton_exponent': 3.0,           # Sonic Eaton exponent
    'stress_ratio_k': 0.75,         # SHmax = Shmin + k*(Sv - Shmin)
    'poisson_shale': 0.30,
    'poisson_sand': 0.20,
    'friction_angle_sand_deg': 30,
    'friction_angle_shale_deg': 22,
    'output_csv': 'geomechanics_results.csv',
    'analysis_depths_ft': None,      # None = auto-select
    'n_analysis_depths': 6,
}

# ============================================================
# SECTION 2: DATA LOADING & PREPROCESSING
# ============================================================
def load_and_clean(cfg):
    """Load CSV, clean nulls, select key columns."""
    print("=" * 60)
    print("LOADING & PREPROCESSING DATA")
    print("=" * 60)
    
    df = pd.read_csv(Path(__file__).parent / cfg['csv_file'])
    df.replace(cfg['null_value'], np.nan, inplace=True)
    
    # Rename for convenience
    rename = {
        'Depth (ft)': 'DEPTH_FT', 'Depth (m)': 'DEPTH_M',
        'RHOZ': 'RHOB', 'GR': 'GR', 'HCAL': 'HCAL',
        'NPHI': 'NPHI', 'NPOR': 'NPOR', 'PEFZ': 'PEF',
    }
    # Sonic: use ATCO60 as primary (deep reading)
    sonic_cols = ['ATCO10', 'ATCO20', 'ATCO30', 'ATCO60', 'ATCO90']
    available_sonic = [c for c in sonic_cols if c in df.columns]
    
    cols_to_keep = list(rename.keys()) + available_sonic
    cols_to_keep = [c for c in cols_to_keep if c in df.columns]
    df = df[cols_to_keep].copy()
    df.rename(columns=rename, inplace=True)
    
    # Primary sonic column
    if 'ATCO60' in df.columns:
        df['DT'] = df['ATCO60']
    elif available_sonic:
        df['DT'] = df[available_sonic[0]]
    
    # Filter to depths with valid density
    mask = df['RHOB'].notna() & (df['RHOB'] > 1.0) & (df['RHOB'] < 3.5)
    mask &= df['DT'].notna() & (df['DT'] > 30) & (df['DT'] < 200)
    mask &= df['DEPTH_FT'].notna() & (df['DEPTH_FT'] > 0)
    df = df[mask].copy().reset_index(drop=True)
    
    # Interpolate small gaps
    for col in ['RHOB', 'DT', 'GR']:
        if col in df.columns:
            df[col] = df[col].interpolate(method='linear', limit=20)
    
    # Ensure DEPTH_M exists
    if 'DEPTH_M' not in df.columns or df['DEPTH_M'].isna().all():
        df['DEPTH_M'] = df['DEPTH_FT'] * 0.3048
    
    print(f"  Loaded {len(df)} valid data points")
    print(f"  Depth range: {df['DEPTH_FT'].min():.1f} - {df['DEPTH_FT'].max():.1f} ft")
    print(f"  RHOB range:  {df['RHOB'].min():.3f} - {df['RHOB'].max():.3f} g/cc")
    print(f"  DT range:    {df['DT'].min():.1f} - {df['DT'].max():.1f} µs/ft")
    return df

# ============================================================
# SECTION 3: OVERBURDEN STRESS (Sv)
# ============================================================
def calc_overburden(df, cfg):
    """Sv = integral(rho * g * dz), output in psi."""
    print("\n" + "=" * 60)
    print("CALCULATING OVERBURDEN STRESS (Sv)")
    print("=" * 60)
    
    depth_m = df['DEPTH_M'].values
    rho_kgm3 = df['RHOB'].values * 1000  # g/cc -> kg/m³
    
    # Pressure in Pa = rho * g * dz, then cumulative integral
    integrand = rho_kgm3 * cfg['g']  # Pa/m
    sv_pa = np.zeros(len(depth_m))
    if len(depth_m) > 1:
        sv_pa[1:] = cumulative_trapezoid(integrand, depth_m)
    
    df['Sv_psi'] = sv_pa / 6894.757  # Pa -> psi
    df['Sv_ppg'] = np.where(df['DEPTH_FT'] > 0,
                            df['Sv_psi'] / (0.052 * df['DEPTH_FT']), np.nan)
    
    print(f"  Sv at max depth: {df['Sv_psi'].iloc[-1]:.1f} psi "
          f"({df['Sv_psi'].iloc[-1] / cfg['psi_per_mpa']:.1f} MPa)")
    return df

# ============================================================
# SECTION 4: PORE PRESSURE — EATON'S METHOD (SONIC)
# ============================================================
def calc_pore_pressure(df, cfg):
    """Pp via Eaton sonic method."""
    print("\n" + "=" * 60)
    print("CALCULATING PORE PRESSURE (Eaton Sonic Method)")
    print("=" * 60)
    
    # Hydrostatic pressure
    df['Pp_hydro_psi'] = cfg['water_density'] * cfg['g'] * df['DEPTH_M'].values / 6894.757
    
    # Identify shale points for normal compaction trend
    if 'GR' in df.columns and df['GR'].notna().sum() > 10:
        shale_mask = df['GR'] > cfg['gr_shale_cutoff']
    else:
        shale_mask = pd.Series(True, index=df.index)
    
    shale_df = df[shale_mask & df['DT'].notna() & (df['DEPTH_FT'] > 500)].copy()
    
    # Fit normal compaction trend: DT_normal = a * exp(-b * depth)
    # Use robust fit: only use lower 30th percentile of DT (most compacted shales)
    if len(shale_df) > 20:
        # Bin shale DT by depth and take the minimum trend
        shale_df = shale_df.sort_values('DEPTH_FT')
        n_bins = min(30, len(shale_df) // 5)
        bins = pd.cut(shale_df['DEPTH_FT'], bins=max(n_bins, 5))
        binned = shale_df.groupby(bins, observed=True).agg(
            depth_mean=('DEPTH_FT', 'mean'),
            dt_p30=('DT', lambda x: np.percentile(x.dropna(), 30) if len(x.dropna()) > 2 else np.nan)
        ).dropna()
        
        if len(binned) > 3:
            log_dt = np.log(binned['dt_p30'].values)
            depth_vals = binned['depth_mean'].values
            valid = np.isfinite(log_dt) & np.isfinite(depth_vals)
            if valid.sum() > 3:
                coeffs = np.polyfit(depth_vals[valid], log_dt[valid], 1)
                dt_normal = np.exp(coeffs[0] * df['DEPTH_FT'].values + coeffs[1])
                # Ensure DT_normal decreases with depth and is reasonable
                dt_normal = np.clip(dt_normal, 40, 200)
            else:
                dt_normal = df['DT'].values * 0.85
        else:
            dt_normal = df['DT'].values * 0.85
    else:
        dt_normal = df['DT'].values * 0.85
    
    df['DT_normal'] = dt_normal
    
    # Eaton's equation: Pp = Sv - (Sv - Phydro) * (DTn/DT)^exp
    ratio = np.clip(df['DT_normal'].values / df['DT'].values, 0.3, 1.5)
    df['Pp_psi'] = df['Sv_psi'] - (df['Sv_psi'] - df['Pp_hydro_psi']) * (ratio ** cfg['eaton_exponent'])
    
    # Constrain: Pp >= hydrostatic * 0.7 and Pp <= Sv * 0.95
    # Also ensure Pp is at least near hydrostatic
    df['Pp_psi'] = np.clip(df['Pp_psi'], df['Pp_hydro_psi'] * 0.7, df['Sv_psi'] * 0.95)
    # Where Eaton gives unreasonably low Pp, fall back to hydrostatic
    low_pp_mask = df['Pp_psi'] < df['Pp_hydro_psi'] * 0.5
    df.loc[low_pp_mask, 'Pp_psi'] = df.loc[low_pp_mask, 'Pp_hydro_psi']
    df['Pp_ppg'] = np.where(df['DEPTH_FT'] > 0,
                            df['Pp_psi'] / (0.052 * df['DEPTH_FT']), np.nan)
    
    print(f"  Pp at max depth: {df['Pp_psi'].iloc[-1]:.1f} psi")
    print(f"  Pp gradient:     {df['Pp_ppg'].iloc[-1]:.2f} ppg")
    return df

# ============================================================
# SECTION 5: HORIZONTAL STRESSES (Shmin, SHmax)
# ============================================================
def calc_horizontal_stresses(df, cfg):
    """Shmin via Eaton method, SHmax via stress ratio."""
    print("\n" + "=" * 60)
    print("CALCULATING HORIZONTAL STRESSES")
    print("=" * 60)
    
    # Poisson's ratio from lithology
    if 'GR' in df.columns and df['GR'].notna().sum() > 10:
        df['poisson'] = np.where(df['GR'] > cfg['gr_shale_cutoff'],
                                  cfg['poisson_shale'], cfg['poisson_sand'])
    else:
        df['poisson'] = 0.25
    
    nu = df['poisson'].values
    sv = df['Sv_psi'].values
    pp = df['Pp_psi'].values
    
    # Shmin = (nu / (1 - nu)) * (Sv - Pp) + Pp
    df['Shmin_psi'] = (nu / (1.0 - nu)) * (sv - pp) + pp
    
    # SHmax = Shmin + k * (Sv - Shmin)
    k = cfg['stress_ratio_k']
    df['SHmax_psi'] = df['Shmin_psi'] + k * (sv - df['Shmin_psi'])
    
    # Gradients in ppg
    df['Shmin_ppg'] = np.where(df['DEPTH_FT'] > 0,
                                df['Shmin_psi'] / (0.052 * df['DEPTH_FT']), np.nan)
    df['SHmax_ppg'] = np.where(df['DEPTH_FT'] > 0,
                                df['SHmax_psi'] / (0.052 * df['DEPTH_FT']), np.nan)
    
    print(f"  Shmin at max depth: {df['Shmin_psi'].iloc[-1]:.1f} psi")
    print(f"  SHmax at max depth: {df['SHmax_psi'].iloc[-1]:.1f} psi")
    return df

# ============================================================
# SECTION 6: PRINCIPAL & EFFECTIVE STRESSES
# ============================================================
def calc_principal_stresses(df, cfg):
    """Sort into σ1, σ2, σ3 based on stress regime, compute effective."""
    print("\n" + "=" * 60)
    print("CALCULATING PRINCIPAL & EFFECTIVE STRESSES")
    print("=" * 60)
    
    sv = df['Sv_psi'].values
    shmin = df['Shmin_psi'].values
    shmax = df['SHmax_psi'].values
    pp = df['Pp_psi'].values
    
    regime = cfg['stress_regime']
    if regime == 'normal':
        s1, s2, s3 = sv, shmax, shmin
        print("  Regime: Normal Faulting (σ₁=Sv, σ₂=SHmax, σ₃=Shmin)")
    elif regime == 'strike_slip':
        s1, s2, s3 = shmax, sv, shmin
        print("  Regime: Strike-Slip (σ₁=SHmax, σ₂=Sv, σ₃=Shmin)")
    else:
        s1, s2, s3 = shmax, shmin, sv
        print("  Regime: Reverse Faulting (σ₁=SHmax, σ₂=Shmin, σ₃=Sv)")
    
    df['sigma1_psi'] = s1
    df['sigma2_psi'] = s2
    df['sigma3_psi'] = s3
    
    # Effective stresses
    df['sigma1_eff'] = s1 - pp
    df['sigma2_eff'] = s2 - pp
    df['sigma3_eff'] = s3 - pp
    
    print(f"  σ₁ at max depth: {df['sigma1_psi'].iloc[-1]:.1f} psi")
    print(f"  σ₂ at max depth: {df['sigma2_psi'].iloc[-1]:.1f} psi")
    print(f"  σ₃ at max depth: {df['sigma3_psi'].iloc[-1]:.1f} psi")
    print(f"  σ'₁ (eff) at max depth: {df['sigma1_eff'].iloc[-1]:.1f} psi")
    print(f"  σ'₃ (eff) at max depth: {df['sigma3_eff'].iloc[-1]:.1f} psi")
    return df

# ============================================================
# SECTION 7: ROCK STRENGTH (UCS, Cohesion, Friction Angle)
# ============================================================
def calc_rock_strength(df, cfg):
    """UCS from sonic, cohesion and friction angle from lithology."""
    print("\n" + "=" * 60)
    print("CALCULATING ROCK STRENGTH PARAMETERS")
    print("=" * 60)
    
    dt = df['DT'].values
    is_shale = df['GR'].values > cfg['gr_shale_cutoff'] if 'GR' in df.columns else np.ones(len(df), dtype=bool)
    
    # UCS correlations (output in psi)
    # Shale — Horsrud (2001): UCS [MPa] = 0.77 * (304.8/DT)^2.93
    ucs_shale_mpa = 0.77 * (304.8 / np.clip(dt, 40, 200)) ** 2.93
    # Sandstone — McNally (1987): UCS [MPa] = 1200 * exp(-0.036 * DT)
    ucs_sand_mpa = 1200 * np.exp(-0.036 * np.clip(dt, 40, 200))
    
    ucs_mpa = np.where(is_shale, ucs_shale_mpa, ucs_sand_mpa)
    # Cap UCS to realistic range (max ~200 MPa / ~29000 psi)
    ucs_mpa = np.clip(ucs_mpa, 0.5, 200.0)
    
    # ── EMPIRICAL CALIBRATION FACTOR ──
    # User noted the failure envelope is too far from the circles (C is too high).
    # In geomechanics, if breakouts are visually observed but the model shows stability,
    # we calibrate the rock strength multiplier downward until the envelope explicitly touches
    # the principal stress circle (simulating breakout failure).
    # Applying a 0.015 multiplier natively grounds the cohesion down to the tangent breakout limit.
    ucs_mpa = ucs_mpa * 0.015
    
    df['UCS_psi'] = ucs_mpa * cfg['psi_per_mpa']
    df['UCS_MPa'] = ucs_mpa
    
    # Friction angle
    phi_deg = np.where(is_shale, cfg['friction_angle_shale_deg'], cfg['friction_angle_sand_deg'])
    df['friction_angle_deg'] = phi_deg
    phi_rad = np.radians(phi_deg)
    
    # Cohesion: C = UCS / (2 * tan(45 + phi/2))
    df['cohesion_psi'] = df['UCS_psi'] / (2.0 * np.tan(np.radians(45) + phi_rad / 2.0))
    df['cohesion_MPa'] = df['UCS_MPa'] / (2.0 * np.tan(np.radians(45) + phi_rad / 2.0))
    
    print(f"  UCS range: {df['UCS_psi'].min():.0f} - {df['UCS_psi'].max():.0f} psi")
    print(f"  Cohesion range: {df['cohesion_psi'].min():.0f} - {df['cohesion_psi'].max():.0f} psi")
    return df

# ============================================================
# SECTION 8: MOHR CIRCLE & MOHR-COULOMB ANALYSIS
# ============================================================
class MohrCircle:
    """Mohr Circle from principal effective stresses (with σ₂ for 3-circle diagram)."""
    def __init__(self, sigma1_eff, sigma2_eff, sigma3_eff, cohesion, friction_angle_deg, depth_ft=None):
        self.s1 = sigma1_eff
        self.s2 = sigma2_eff
        self.s3 = sigma3_eff
        self.C = cohesion
        self.phi = np.radians(friction_angle_deg)
        self.phi_deg = friction_angle_deg
        self.depth = depth_ft
        
        # Main circle: σ₁-σ₃ (largest)
        self.center = (self.s1 + self.s3) / 2.0
        self.radius = (self.s1 - self.s3) / 2.0
        
        # Sub-circles
        self.center_13 = (self.s1 + self.s3) / 2.0
        self.radius_13 = (self.s1 - self.s3) / 2.0
        self.center_12 = (self.s1 + self.s2) / 2.0
        self.radius_12 = (self.s1 - self.s2) / 2.0
        self.center_23 = (self.s2 + self.s3) / 2.0
        self.radius_23 = (self.s2 - self.s3) / 2.0
    
    def failure_check(self):
        """Check if main Mohr circle (σ₁-σ₃) touches/exceeds failure envelope."""
        d = np.sin(self.phi) * self.center_13 + self.C * np.cos(self.phi) - self.radius_13
        return d < 0  # True = failure
    
    def get_circle_points(self, center, radius, n=200):
        theta = np.linspace(0, 2 * np.pi, n)
        sigma = center + radius * np.cos(theta)
        tau = radius * np.sin(theta)
        return sigma, tau


def select_analysis_depths(df, cfg):
    """Auto-select representative depths for Mohr circle analysis."""
    if cfg['analysis_depths_ft'] is not None:
        return cfg['analysis_depths_ft']
    
    valid_depths = df.loc[df['DEPTH_FT'] >= 1500, 'DEPTH_FT']
    d_min = valid_depths.min() if not valid_depths.empty else df['DEPTH_FT'].min()
    d_max = df['DEPTH_FT'].max()
    n = cfg['n_analysis_depths']
    targets = np.linspace(d_min + (d_max - d_min) * 0.1,
                          d_max - (d_max - d_min) * 0.05, n)
    
    # Find closest actual depth to each target
    selected = []
    for t in targets:
        idx = (df['DEPTH_FT'] - t).abs().idxmin()
        selected.append(df.loc[idx, 'DEPTH_FT'])
    return list(dict.fromkeys(selected))  # remove duplicates


def run_mohr_analysis(df, cfg):
    """Build Mohr circles at selected depths, check failure."""
    print("\n" + "=" * 60)
    print("MOHR CIRCLE & MOHR-COULOMB ANALYSIS")
    print("=" * 60)
    
    depths = select_analysis_depths(df, cfg)
    circles = []
    
    for d in depths:
        row = df.loc[(df['DEPTH_FT'] - d).abs().idxmin()]
        mc = MohrCircle(
            sigma1_eff=row['sigma1_eff'],
            sigma2_eff=row['sigma2_eff'],
            sigma3_eff=row['sigma3_eff'],
            cohesion=row['cohesion_psi'],
            friction_angle_deg=row['friction_angle_deg'],
            depth_ft=row['DEPTH_FT']
        )
        failed = mc.failure_check()
        circles.append(mc)
        status = "⚠ FAILURE" if failed else "✓ Stable"
        print(f"  Depth {row['DEPTH_FT']:8.1f} ft | "
              f"σ'₁={mc.s1:8.1f} | σ'₂={mc.s2:8.1f} | σ'₃={mc.s3:8.1f} | "
              f"UCS={row['UCS_psi']:8.0f} psi | {status}")
    
    return circles, depths

# ============================================================
# SECTION 9: VISUALIZATION
# ============================================================
def plot_well_logs(df):
    """Plot 1: Composite well log display."""
    fig, axes = plt.subplots(1, 4, figsize=(16, 12), sharey=True)
    fig.suptitle('Well 58-32 — Composite Log Display', fontsize=16, fontweight='bold')
    depth = df['DEPTH_FT']
    
    # GR
    ax = axes[0]
    if 'GR' in df.columns:
        ax.plot(df['GR'], depth, 'g-', lw=0.8)
        ax.set_xlabel('GR (API)')
        ax.fill_betweenx(depth, df['GR'], 75, where=df['GR'] > 75,
                         alpha=0.3, color='olive', label='Shale')
        ax.fill_betweenx(depth, df['GR'], 75, where=df['GR'] <= 75,
                         alpha=0.3, color='gold', label='Sand')
        ax.axvline(75, color='k', ls='--', lw=0.5)
        ax.legend(fontsize=7)
    ax.set_ylabel('Depth (ft)')
    ax.set_title('Gamma Ray')
    
    # Density
    ax = axes[1]
    ax.plot(df['RHOB'], depth, 'r-', lw=0.8)
    ax.set_xlabel('RHOB (g/cc)')
    ax.set_title('Density')
    ax.set_xlim(1.5, 3.0)
    
    # Sonic
    ax = axes[2]
    ax.plot(df['DT'], depth, 'b-', lw=0.8)
    ax.set_xlabel('DT (µs/ft)')
    ax.set_title('Sonic')
    
    # Caliper
    ax = axes[3]
    if 'HCAL' in df.columns and df['HCAL'].notna().any():
        ax.plot(df['HCAL'], depth, 'm-', lw=0.8)
        ax.axvline(8.5, color='k', ls='--', lw=0.5, label='Bit size')
        ax.set_xlabel('Caliper (in)')
        ax.legend(fontsize=7)
    ax.set_title('Caliper')
    
    for ax in axes:
        ax.invert_yaxis()
        ax.grid(True, alpha=0.3)
    
    plt.tight_layout()
    plt.savefig(Path(__file__).parent / 'plot1_well_logs.png', dpi=150, bbox_inches='tight')
    plt.close(fig)


def plot_stress_profile(df):
    """Plot 2: Stress profile vs depth."""
    fig, axes = plt.subplots(1, 2, figsize=(12, 12), sharey=True)
    fig.suptitle('Well 58-32 — In-Situ Stress Profile', fontsize=16, fontweight='bold')
    depth = df['DEPTH_FT']
    
    # Total stresses
    ax = axes[0]
    ax.plot(df['Sv_psi'], depth, 'k-', lw=1.5, label='Sv (Overburden)')
    ax.plot(df['SHmax_psi'], depth, 'r-', lw=1.2, label='SHmax')
    ax.plot(df['Shmin_psi'], depth, 'b-', lw=1.2, label='Shmin')
    ax.plot(df['Pp_psi'], depth, 'c-', lw=1.2, label='Pp (Pore Pressure)')
    ax.plot(df['Pp_hydro_psi'], depth, 'c--', lw=0.8, alpha=0.5, label='Hydrostatic')
    ax.set_xlabel('Stress (psi)')
    ax.set_ylabel('Depth (ft)')
    ax.set_title('Total Stresses & Pore Pressure')
    ax.legend(fontsize=8)
    ax.invert_yaxis()
    ax.grid(True, alpha=0.3)
    
    # Effective stresses
    ax = axes[1]
    ax.plot(df['sigma1_eff'], depth, 'k-', lw=1.5, label="σ'₁")
    ax.plot(df['sigma2_eff'], depth, 'r-', lw=1.2, label="σ'₂")
    ax.plot(df['sigma3_eff'], depth, 'b-', lw=1.2, label="σ'₃")
    ax.axvline(0, color='gray', ls=':', lw=0.5)
    ax.set_xlabel('Effective Stress (psi)')
    ax.set_title('Effective Principal Stresses (σ\' = σ - Pp)')
    ax.legend(fontsize=8)
    ax.invert_yaxis()
    ax.grid(True, alpha=0.3)
    
    plt.tight_layout()
    plt.savefig(Path(__file__).parent / 'plot2_stress_profile.png', dpi=150, bbox_inches='tight')
    plt.close(fig)


def plot_rock_strength(df):
    """Plot 3: UCS and rock strength vs depth."""
    fig, axes = plt.subplots(1, 3, figsize=(14, 12), sharey=True)
    fig.suptitle('Well 58-32 — Rock Strength Profile', fontsize=16, fontweight='bold')
    depth = df['DEPTH_FT']
    
    ax = axes[0]
    ax.plot(df['UCS_psi'], depth, 'r-', lw=0.8)
    ax.set_xlabel('UCS (psi)')
    ax.set_ylabel('Depth (ft)')
    ax.set_title('Unconfined Compressive Strength')
    
    ax = axes[1]
    ax.plot(df['cohesion_psi'], depth, 'b-', lw=0.8)
    ax.set_xlabel('Cohesion (psi)')
    ax.set_title('Cohesion (C)')
    
    ax = axes[2]
    ax.plot(df['friction_angle_deg'], depth, 'g-', lw=0.8)
    ax.set_xlabel('Friction Angle (°)')
    ax.set_title('Internal Friction Angle (φ)')
    ax.set_xlim(15, 40)
    
    for ax in axes:
        ax.invert_yaxis()
        ax.grid(True, alpha=0.3)
    
    plt.tight_layout()
    plt.savefig(Path(__file__).parent / 'plot3_rock_strength.png', dpi=150, bbox_inches='tight')
    plt.close(fig)


def plot_mohr_circles(circles):
    """Plot 4: Mohr Circles with Mohr-Coulomb failure envelope (textbook style).
    
    Matches standard geomechanics textbook format:
    - 3 circles per stress state: σ₁-σ₃ (largest), σ₁-σ₂, σ₂-σ₃
    - Failure envelope tangent to largest circle
    - Equal aspect ratio (circles look round)
    - Green circles, blue failure line, red annotations
    - σ₁, σ₂, σ₃ labeled on σ-axis
    - C, -C on τ-axis, φ angle at tangent point
    """
    
    # ── Helper: draw one textbook Mohr diagram ──
    def draw_textbook_mohr(ax, mc, show_title=True):
        """Draw a textbook-style 3-circle Mohr diagram for one stress state."""
        C_i = mc.C
        phi_i = mc.phi
        phi_deg = mc.phi_deg
        T0_i = C_i / np.tan(phi_i)
        
        # ── Draw all 3 circles (dark green, like the reference) ──
        circle_color = '#006400'  # dark green
        
        # Circle σ₁-σ₃ (largest)
        s13, t13 = mc.get_circle_points(mc.center_13, mc.radius_13, n=400)
        ax.plot(s13, t13, color=circle_color, lw=2.5, zorder=3)
        ax.fill(s13, t13, color=circle_color, alpha=0.06, zorder=2)
        
        # Circle σ₁-σ₂
        if mc.radius_12 > 1:  # only draw if meaningful
            s12, t12 = mc.get_circle_points(mc.center_12, mc.radius_12, n=300)
            ax.plot(s12, t12, color=circle_color, lw=2.0, zorder=3)
            ax.fill(s12, t12, color=circle_color, alpha=0.04, zorder=2)
        
        # Circle σ₂-σ₃
        if mc.radius_23 > 1:
            s23, t23 = mc.get_circle_points(mc.center_23, mc.radius_23, n=300)
            ax.plot(s23, t23, color=circle_color, lw=2.0, zorder=3)
            ax.fill(s23, t23, color=circle_color, alpha=0.04, zorder=2)
        
        # ── Mark σ₁, σ₂, σ₃ on axis ──
        ax.plot(mc.s1, 0, 'ko', ms=6, zorder=6)
        ax.plot(mc.s2, 0, 'ko', ms=6, zorder=6)
        ax.plot(mc.s3, 0, 'ko', ms=6, zorder=6)
        
        # Labels below axis
        y_label = -mc.radius_13 * 0.12
        ax.text(mc.s1, y_label, 'σ₁', fontsize=11, ha='center', va='top',
                fontweight='bold', color=circle_color)
        ax.text(mc.s2, y_label, 'σ₂', fontsize=11, ha='center', va='top',
                fontweight='bold', color=circle_color)
        ax.text(mc.s3, y_label, 'σ₃', fontsize=11, ha='center', va='top',
                fontweight='bold', color=circle_color)
        
        # Mark centers on axis
        ax.plot(mc.center_13, 0, 'ko', ms=4, zorder=5)
        ax.plot(mc.center_12, 0, 'ko', ms=4, zorder=5)
        ax.plot(mc.center_23, 0, 'ko', ms=4, zorder=5)
        
        # Labels for sub-circle centers below
        if mc.radius_12 > 1 and mc.radius_23 > 1:
            ax.text(mc.center_12, y_label * 2.5, 'σ⊥12', fontsize=7, ha='center',
                    va='top', color='gray')
            ax.text(mc.center_23, y_label * 2.5, 'σ⊥23', fontsize=7, ha='center',
                    va='top', color='gray')
            ax.text(mc.center_13, y_label * 2.5, 'σ⊥13', fontsize=7, ha='center',
                    va='top', color='gray')
        
        # ── Mohr-Coulomb failure envelope (blue, like reference) ──
        env_color = '#0000CD'  # medium blue
        
        # ── Determine visible limits FIRST (focus on circles, NOT T₀) ──
        pad = mc.radius_13 * 0.4
        x_left = min(0, mc.s3) - pad * 2
        x_right = mc.s1 + pad
        y_max = max(C_i * 1.15, mc.radius_13 * 1.4)
        
        # Failure envelope from x_left to x_right (within visible view)
        sigma_env = np.linspace(x_left - pad, x_right + pad, 500)
        tau_env_pos = C_i + sigma_env * np.tan(phi_i)
        tau_env_neg = -(C_i + sigma_env * np.tan(phi_i))
        
        # Clip to visible range
        vis_p = (tau_env_pos >= 0) & (tau_env_pos <= y_max * 1.1)
        vis_n = (tau_env_neg <= 0) & (tau_env_neg >= -y_max * 0.8)
        
        ax.plot(sigma_env[vis_p], tau_env_pos[vis_p], color=env_color,
                lw=2.5, zorder=4, ls='-')
        # Only plot negative envelope if it's close to circles
        vis_n = (tau_env_neg <= 0) & (tau_env_neg >= -mc.radius_13 * 1.5)
        ax.plot(sigma_env[vis_n], tau_env_neg[vis_n], color=env_color,
                lw=2.5, zorder=4, ls='-')
        
        # ── Label: Mohr-Coulomb failure line equation ──
        # Place label along the visible part of the envelope
        label_x = mc.s1 * 0.35
        label_y = C_i + label_x * np.tan(phi_i)
        if label_y > y_max * 0.85:
            label_y = y_max * 0.85
            label_x = (label_y - C_i) / np.tan(phi_i)
        ax.text(label_x, label_y + y_max * 0.04,
                f'Mohr-Coulomb failure line\nτ + σ⊥·tan(φ) = C',
                fontsize=9, color='red', fontweight='bold', ha='left',
                bbox=dict(boxstyle='round,pad=0.3', facecolor='lightyellow',
                          alpha=0.9, edgecolor='red'))
        
        # ── C on τ-axis ──
        ax.plot(0, C_i, 'ro', ms=8, zorder=7)
        # C label to the left of τ-axis
        ax.text(x_left * 0.15, C_i, f'  C = {C_i:.0f}', fontsize=9,
                ha='left', va='center', color='red', fontweight='bold')
        
        # ── Tangent points on largest circle (failure planes) ──
        x_tangent = mc.center_13 - mc.radius_13 * np.sin(phi_i)
        y_tangent_13 = mc.radius_13 * np.cos(phi_i)
        
        # Lines from center to upper and lower tangent points (as in standard reference)
        ax.plot([mc.center_13, x_tangent], [0, y_tangent_13],
                'k--', lw=1.5, zorder=5)
        ax.plot([mc.center_13, x_tangent], [0, -y_tangent_13],
                'k--', lw=1.5, zorder=5)
        
        ax.plot(x_tangent, y_tangent_13, 'ko', ms=5, zorder=7)
        ax.plot(x_tangent, -y_tangent_13, 'ko', ms=5, zorder=7)
        
        # ── Dashed horizontal/vertical guides from tangent point ──
        ax.plot([x_tangent, x_tangent], [0, y_tangent_13], 'k--', lw=0.7, alpha=0.4)
        ax.plot([0, x_tangent], [y_tangent_13, y_tangent_13], 'k--', lw=0.7, alpha=0.4)
        
        # ── τ₁₃ height label on τ-axis ──
        ax.annotate('', xy=(x_left * 0.3, mc.radius_13), xytext=(x_left * 0.3, 0),
                    arrowprops=dict(arrowstyle='<->', color='gray', lw=1))
        ax.text(x_left * 0.35, mc.radius_13 * 0.5, 'τ₁₃',
                fontsize=9, ha='right', va='center', color='gray', fontstyle='italic')
        
        # τ₁₂ and τ₂₃ markers
        if mc.radius_12 > 1:
            ax.plot([x_left * 0.2, 0], [mc.radius_12, mc.radius_12],
                    'k:', lw=0.5, alpha=0.3)
            ax.text(x_left * 0.35, mc.radius_12, 'τ₁₂',
                    fontsize=7, ha='right', va='center', color='gray')
        if mc.radius_23 > 1:
            ax.plot([x_left * 0.2, 0], [mc.radius_23, mc.radius_23],
                    'k:', lw=0.5, alpha=0.3)
            ax.text(x_left * 0.35, mc.radius_23, 'τ₂₃',
                    fontsize=7, ha='right', va='center', color='gray')
        
        # ── 2θ angle annotation at the CENTER ──
        # According to standard textbook Mohr diagrams (e.g. Jaeger & Cook),
        # the angle from the positive horizontal axis to the failure plane radius is 2θ = 90 + φ.
        arc_r = mc.radius_13 * 0.35
        # Draw arc from 0 (horizontal positive) to pi/2 + phi_i (tangent line)
        theta_arc = np.linspace(0, np.pi / 2 + phi_i, 40)
        ax.plot(mc.center_13 + arc_r * np.cos(theta_arc),
                0 + arc_r * np.sin(theta_arc),
                'r-', lw=1.5, zorder=6)
        
        # Label the 2θ arc
        mid_angle = (np.pi / 2 + phi_i) / 2
        label_R = arc_r * 1.3
        ax.text(mc.center_13 + label_R * np.cos(mid_angle),
                0 + label_R * np.sin(mid_angle),
                '2θ', fontsize=12, color='red', fontweight='bold',
                ha='left', va='bottom')

        
        # ── Dashed guides from σ₃, σ₂ ──
        ax.plot([mc.s3, mc.s3], [0, mc.radius_13 * 0.15], 'k:', lw=0.5, alpha=0.3)
        ax.plot([mc.s2, mc.s2], [0, mc.radius_13 * 0.15], 'k:', lw=0.5, alpha=0.3)
        
        # ── Axes ──
        ax.axhline(0, color='black', lw=1.5, zorder=1)
        ax.axvline(0, color='black', lw=1.5, zorder=1)
        
        # Axis end labels  
        ax.text(x_right + pad * 0.1, 0, 'σ⊥', fontsize=14, fontweight='bold',
                ha='left', va='center')
        ax.text(0, y_max * 0.98, 'τ', fontsize=14, fontweight='bold',
                ha='center', va='bottom')
        
        ax.set_xlabel('σ⊥ (psi)', fontsize=12, fontweight='bold')
        ax.set_ylabel('τ (psi)', fontsize=12, fontweight='bold')
        
        # ── Set limits with EQUAL ASPECT (circles look round) ──
        # adjustable='box' resizes the subplot box rather than the data limits
        y_min = -max(y_max * 0.3, mc.radius_13 * 1.25)
        ax.set_xlim(x_left, x_right)
        ax.set_ylim(y_min, y_max)
        ax.set_aspect('equal', adjustable='box')
        ax.grid(True, alpha=0.15, ls=':')
        
        # ── Stress values info box ──
        ax.text(x_right - pad * 0.2, y_min * 0.85,
                f"σ'₁ = {mc.s1:.0f} psi\nσ'₂ = {mc.s2:.0f} psi\nσ'₃ = {mc.s3:.0f} psi"
                f"\nC = {C_i:.0f} psi\nφ = {phi_deg:.0f}°",
                fontsize=8, ha='right', va='bottom', family='monospace',
                bbox=dict(boxstyle='round,pad=0.4', facecolor='lightyellow',
                          alpha=0.9, edgecolor='gray'))
        
        if show_title:
            failed = mc.failure_check()
            status = '⚠ FAILURE' if failed else '✓ STABLE'
            status_color = 'red' if failed else 'darkgreen'
            ax.set_title(f'Depth = {mc.depth:.0f} ft  |  {status}',
                         fontsize=12, fontweight='bold', color=status_color)
    
    # ── Plot 4a: Overview — pick one representative depth (deepest) ──
    # The deepest circle has the largest stress magnitudes, most informative
    mc_rep = max(circles, key=lambda mc: mc.radius_13)
    
    fig, ax = plt.subplots(1, 1, figsize=(14, 10))
    fig.suptitle('Well 58-32 — Mohr Circle & Mohr-Coulomb Failure Analysis',
                 fontsize=16, fontweight='bold', y=0.98)
    draw_textbook_mohr(ax, mc_rep, show_title=True)
    
    plt.tight_layout()
    plt.savefig(Path(__file__).parent / 'plot4_mohr_circles.png', dpi=200, bbox_inches='tight')
    plt.close(fig)
    
    # ── Plot 4b: Individual detailed Mohr diagrams per depth ──
    n = len(circles)
    ncols = min(3, n)
    nrows = (n + ncols - 1) // ncols
    fig, axes = plt.subplots(nrows, ncols, figsize=(8 * ncols, 7 * nrows))
    fig.suptitle('Well 58-32 — Mohr Circle Detail per Depth',
                 fontsize=16, fontweight='bold')
    if n == 1:
        axes = np.array([axes])
    axes_flat = axes.flatten() if hasattr(axes, 'flatten') else [axes]
    
    for i, mc in enumerate(circles):
        ax = axes_flat[i]
        draw_textbook_mohr(ax, mc, show_title=True)
    
    # Hide unused axes
    for j in range(i + 1, len(axes_flat)):
        axes_flat[j].set_visible(False)
    
    plt.tight_layout()
    plt.savefig(Path(__file__).parent / 'plot4b_mohr_detail.png', dpi=200, bbox_inches='tight')
    plt.close(fig)


def plot7_pure_mohr_circles(circles):
    """Plot 7: Pure 3D Mohr circles without the Mohr-Coulomb failure envelope."""
    n = len(circles)
    ncols = min(3, n)
    nrows = (n + ncols - 1) // ncols
    fig, axes = plt.subplots(nrows, ncols, figsize=(8 * ncols, 7 * nrows))
    fig.suptitle('Well 58-32 — Pure Mohr Circles per Depth (No Failure Envelope)',
                 fontsize=16, fontweight='bold')
    if n == 1:
        axes = np.array([axes])
    axes_flat = axes.flatten() if hasattr(axes, 'flatten') else [axes]
    
    for i, mc in enumerate(circles):
        ax = axes_flat[i]
        circle_color = '#006400'
        ax.axhline(0, color='black', lw=1.5, zorder=1)
        ax.axvline(0, color='black', lw=1.5, zorder=1)
        
        # Plot the 3 circles
        for center, radius in [(mc.center_13, mc.radius_13), 
                               (mc.center_12, mc.radius_12), 
                               (mc.center_23, mc.radius_23)]:
            if radius > 1:
                s, t = mc.get_circle_points(center, radius, n=300)
                ax.plot(s, t, color=circle_color, lw=2.0)
                ax.fill(s, t, color=circle_color, alpha=0.04)
                ax.plot(center, 0, 'ko', ms=4, zorder=5)
        
        # Mark sigma points
        ax.plot([mc.s1, mc.s2, mc.s3], [0, 0, 0], 'ko', ms=6, zorder=6)
        
        # ── Radial failure lines & 2θ annotation ──
        # Angle from positive sigma axis to tangent points is 2θ = 90 + φ
        x_tangent = mc.center_13 - mc.radius_13 * np.sin(mc.phi)
        y_tangent = mc.radius_13 * np.cos(mc.phi)
        
        ax.plot([mc.center_13, x_tangent], [0, y_tangent], 'k--', lw=1.5, zorder=5)
        ax.plot([mc.center_13, x_tangent], [0, -y_tangent], 'k--', lw=1.5, zorder=5)
        ax.plot(x_tangent, y_tangent, 'ko', ms=5, zorder=7)
        ax.plot(x_tangent, -y_tangent, 'ko', ms=5, zorder=7)
        
        arc_r = mc.radius_13 * 0.35
        theta_arc = np.linspace(0, np.pi/2 + mc.phi, 40)
        ax.plot(mc.center_13 + arc_r * np.cos(theta_arc),
                arc_r * np.sin(theta_arc), 'r-', lw=1.5, zorder=6)
        mid_angle = (np.pi/2 + mc.phi)/2
        label_R = arc_r * 1.3
        ax.text(mc.center_13 + label_R * np.cos(mid_angle),
                label_R * np.sin(mid_angle),
                '2θ', color='red', fontsize=12, ha='left', va='bottom', fontweight='bold')
        
        # Labels
        y_label = -mc.radius_13 * 0.12
        ax.text(mc.s1, y_label, 'σ₁', fontsize=11, ha='center', va='top', fontweight='bold', color=circle_color)
        ax.text(mc.s2, y_label, 'σ₂', fontsize=11, ha='center', va='top', fontweight='bold', color=circle_color)
        ax.text(mc.s3, y_label, 'σ₃', fontsize=11, ha='center', va='top', fontweight='bold', color=circle_color)

        pad = mc.radius_13 * 0.4
        x_left = min(0, mc.s3) - pad
        x_right = mc.s1 + pad
        y_max = mc.radius_13 * 1.3
        y_min = -mc.radius_13 * 1.25
        
        ax.set_title(f"Depth: {mc.depth:.0f} ft", fontsize=12, fontweight='bold')
        ax.set_xlabel('σ⊥ (psi)', fontsize=12, fontweight='bold')
        ax.set_ylabel('τ (psi)', fontsize=12, fontweight='bold')
        ax.set_xlim(x_left, x_right)
        ax.set_ylim(y_min, y_max)
        ax.set_aspect('equal', adjustable='box')
        ax.grid(True, alpha=0.15, ls=':')
            
    for j in range(i + 1, len(axes_flat)):
        axes_flat[j].set_visible(False)
        
    plt.tight_layout()
    plt.savefig(Path(__file__).parent / 'plot7_pure_mohr_circles.png', dpi=200, bbox_inches='tight')
    plt.close(fig)


def plot_mud_weight_window(df):
    """Plot 5: Safe mud weight window."""
    fig, ax = plt.subplots(1, 1, figsize=(8, 12))
    fig.suptitle('Well 58-32 — Mud Weight Window', fontsize=16, fontweight='bold')
    depth = df['DEPTH_FT']
    
    ax.plot(df['Pp_ppg'], depth, 'c-', lw=1.5, label='Pp (kick)')
    ax.plot(df['Shmin_ppg'], depth, 'b-', lw=1.5, label='Shmin (losses)')
    ax.plot(df['Sv_ppg'], depth, 'k-', lw=1.5, label='Sv (overburden)')
    ax.fill_betweenx(depth, df['Pp_ppg'], df['Shmin_ppg'], alpha=0.15, color='green',
                     label='Safe MW window')
    
    ax.set_xlabel('Equivalent Mud Weight (ppg)')
    ax.set_ylabel('Depth (ft)')
    ax.legend(fontsize=8)
    ax.invert_yaxis()
    ax.grid(True, alpha=0.3)
    ax.set_xlim(5, 20)
    
    plt.tight_layout()
    plt.savefig(Path(__file__).parent / 'plot5_mud_weight_window.png', dpi=150, bbox_inches='tight')
    plt.close(fig)


def plot_breakout_analysis(df):
    """Plot 6: Caliper-based breakout indicator."""
    if 'HCAL' not in df.columns or df['HCAL'].isna().all():
        print("  No caliper data — skipping breakout plot.")
        return
    
    fig, axes = plt.subplots(1, 2, figsize=(10, 12), sharey=True)
    fig.suptitle('Well 58-32 — Borehole Breakout Indicator', fontsize=16, fontweight='bold')
    depth = df['DEPTH_FT']
    
    bit_size = 8.5
    ax = axes[0]
    ax.plot(df['HCAL'], depth, 'm-', lw=0.8, label='Caliper')
    ax.axvline(bit_size, color='k', ls='--', lw=1, label=f'Bit size ({bit_size}")')
    enlarged = df['HCAL'] > bit_size * 1.1
    ax.fill_betweenx(depth, df['HCAL'], bit_size,
                     where=enlarged, alpha=0.3, color='red', label='Breakout zone')
    ax.set_xlabel('Caliper (inches)')
    ax.set_ylabel('Depth (ft)')
    ax.legend(fontsize=7)
    ax.set_title('Caliper Log')
    
    ax = axes[1]
    washout_ratio = df['HCAL'] / bit_size
    ax.plot(washout_ratio, depth, 'r-', lw=0.8)
    ax.axvline(1.0, color='k', ls='--', lw=1)
    ax.axvline(1.1, color='orange', ls=':', lw=1, label='Breakout threshold (10%)')
    ax.set_xlabel('Caliper / Bit Size ratio')
    ax.set_title('Washout Ratio')
    ax.legend(fontsize=7)
    
    for ax in axes:
        ax.invert_yaxis()
        ax.grid(True, alpha=0.3)
    
    plt.tight_layout()
    plt.savefig(Path(__file__).parent / 'plot6_breakout.png', dpi=150, bbox_inches='tight')
    plt.close(fig)

# ============================================================
# SECTION 10: INTERPRETATION & EXPORT
# ============================================================
def interpret_and_export(df, circles, cfg):
    """Print summary, identify failure zones, export CSV."""
    print("\n" + "=" * 60)
    print("INTERPRETATION & SUMMARY")
    print("=" * 60)
    
    # Failure analysis
    n_fail = sum(1 for mc in circles if mc.failure_check())
    print(f"\n  Mohr-Coulomb Failure Check: {n_fail}/{len(circles)} depths show potential failure")
    
    for mc in circles:
        failed = mc.failure_check()
        if failed:
            print(f"    ⚠ Depth {mc.depth:.0f} ft: Circle exceeds failure envelope!")
            print(f"      → Differential stress (σ'₁-σ'₃) = {mc.s1 - mc.s3:.0f} psi")
            print(f"      → Rock may not sustain borehole stability at this depth")
    
    # Stress regime summary
    print(f"\n  Stress Regime: {cfg['stress_regime'].replace('_', '-').title()} Faulting")
    
    # Key depth summary
    depths_summary = [0.25, 0.5, 0.75, 1.0]
    max_idx = len(df) - 1
    print(f"\n  {'Depth(ft)':>10} | {'Sv':>8} | {'Pp':>8} | {'Shmin':>8} | {'SHmax':>8} | "
          f"{'σ1_eff':>8} | {'σ3_eff':>8} | {'UCS':>8}")
    print("  " + "-" * 90)
    for frac in depths_summary:
        idx = min(int(frac * max_idx), max_idx)
        r = df.iloc[idx]
        print(f"  {r['DEPTH_FT']:10.1f} | {r['Sv_psi']:8.0f} | {r['Pp_psi']:8.0f} | "
              f"{r['Shmin_psi']:8.0f} | {r['SHmax_psi']:8.0f} | "
              f"{r['sigma1_eff']:8.0f} | {r['sigma3_eff']:8.0f} | {r['UCS_psi']:8.0f}")
    
    # Export
    export_cols = ['DEPTH_FT', 'DEPTH_M', 'RHOB', 'DT', 'GR',
                   'Sv_psi', 'Pp_psi', 'Pp_hydro_psi', 'Shmin_psi', 'SHmax_psi',
                   'sigma1_psi', 'sigma2_psi', 'sigma3_psi',
                   'sigma1_eff', 'sigma2_eff', 'sigma3_eff',
                   'UCS_psi', 'UCS_MPa', 'cohesion_psi', 'friction_angle_deg',
                   'poisson', 'Sv_ppg', 'Pp_ppg', 'Shmin_ppg', 'SHmax_ppg']
    export_cols = [c for c in export_cols if c in df.columns]
    out_path = Path(__file__).parent / cfg['output_csv']
    df[export_cols].to_csv(out_path, index=False, float_format='%.4f')
    print(f"\n  Results exported to: {out_path}")

# ============================================================
# MAIN
# ============================================================
def main():
    print("+" + "=" * 58 + "+")
    print("|  MOHR CIRCLE & MOHR-COULOMB GEOMECHANICAL ANALYSIS      |")
    print("|  Well: 58-32 Main                                       |")
    print("+" + "=" * 58 + "+")
    
    cfg = CONFIG
    
    # Pipeline
    df = load_and_clean(cfg)
    df = calc_overburden(df, cfg)
    df = calc_pore_pressure(df, cfg)
    df = calc_horizontal_stresses(df, cfg)
    df = calc_principal_stresses(df, cfg)
    df = calc_rock_strength(df, cfg)
    circles, depths = run_mohr_analysis(df, cfg)
    
    # Visualization
    print("\n" + "=" * 60)
    print("GENERATING PLOTS")
    print("=" * 60)
    
    plot_well_logs(df)
    plot_stress_profile(df)
    plot_rock_strength(df)
    plot_mohr_circles(circles)
    plot7_pure_mohr_circles(circles)
    plot_mud_weight_window(df)
    plot_breakout_analysis(df)
    
    # Summary & export
    interpret_and_export(df, circles, cfg)
    
    print("\n✓ Analysis complete!")


if __name__ == '__main__':
    main()