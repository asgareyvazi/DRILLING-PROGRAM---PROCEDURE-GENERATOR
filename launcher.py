# ============================================================================
# DRILLING PROGRAM & PROCEDURE GENERATOR - PROFESSIONAL EDITION
# Version 3.0
# File: launcher.py (Part 5 of 5)
# Final Integration, Enhanced Word Generator, Setup & Documentation
# ============================================================================

import sys
import os
import subprocess
import json
import shutil
from pathlib import Path
from datetime import datetime


# ============================================================================
# PROJECT STRUCTURE CREATOR
# ============================================================================

class ProjectStructure:
    """ایجاد ساختار پروژه"""

    @staticmethod
    def create_project_structure():
        """ایجاد تمام فولدرها و فایل‌های مورد نیاز"""
        directories = [
            "projects",
            "projects/templates",
            "projects/exports",
            "projects/backup",
            "logs",
            "config",
            "resources",
            "resources/logos",
            "resources/icons",
            "temp",
        ]

        for d in directories:
            Path(d).mkdir(parents=True, exist_ok=True)

        print("✅ Project structure created successfully")
        return True


# ============================================================================
# DEPENDENCY CHECKER & INSTALLER
# ============================================================================

class DependencyManager:
    """مدیریت وابستگی‌ها"""

    REQUIRED_PACKAGES = {
        'PySide6': 'PySide6',
        'docx': 'python-docx',
    }

    OPTIONAL_PACKAGES = {
        'openpyxl': 'openpyxl',
        'matplotlib': 'matplotlib',
        'numpy': 'numpy',
    }

    @staticmethod
    def check_dependencies() -> dict:
        """بررسی وابستگی‌ها"""
        results = {
            'required': {},
            'optional': {},
            'all_required_met': True
        }

        for module, package in DependencyManager.REQUIRED_PACKAGES.items():
            try:
                __import__(module)
                results['required'][package] = True
            except ImportError:
                results['required'][package] = False
                results['all_required_met'] = False

        for module, package in DependencyManager.OPTIONAL_PACKAGES.items():
            try:
                __import__(module)
                results['optional'][package] = True
            except ImportError:
                results['optional'][package] = False

        return results

    @staticmethod
    def install_dependencies(include_optional: bool = False):
        """نصب وابستگی‌ها"""
        packages = list(DependencyManager.REQUIRED_PACKAGES.values())
        if include_optional:
            packages.extend(DependencyManager.OPTIONAL_PACKAGES.values())

        for package in packages:
            try:
                print(f"📦 Installing {package}...")
                subprocess.check_call([
                    sys.executable, "-m", "pip", "install",
                    package, "--quiet"
                ])
                print(f"   ✅ {package} installed successfully")
            except subprocess.CalledProcessError:
                print(f"   ❌ Failed to install {package}")

    @staticmethod
    def print_status():
        """نمایش وضعیت وابستگی‌ها"""
        results = DependencyManager.check_dependencies()

        print("\n" + "=" * 60)
        print("  DEPENDENCY STATUS CHECK")
        print("=" * 60)

        print("\n  Required Packages:")
        for pkg, status in results['required'].items():
            icon = "✅" if status else "❌"
            print(f"    {icon} {pkg}")

        print("\n  Optional Packages:")
        for pkg, status in results['optional'].items():
            icon = "✅" if status else "⚠️"
            print(f"    {icon} {pkg}")

        if results['all_required_met']:
            print("\n  ✅ All required dependencies are met!")
        else:
            print("\n  ❌ Some required dependencies are missing!")
            print("  Run: python launcher.py --install")

        print("=" * 60 + "\n")
        return results['all_required_met']


# ============================================================================
# ENHANCED WORD GENERATOR (COMPLETE INTEGRATION)
# ============================================================================

class EnhancedWordGenerator:
    """تولیدکننده Word پیشرفته با یکپارچه‌سازی تمام ماژول‌ها"""

    def __init__(self, project):
        self.project = project

    def generate_complete_document(self, file_path: str, progress=None):
        """تولید سند کامل با تمام ضمایم"""
        try:
            from word_generator import DrillingProgramWordGenerator
            from advanced_modules import AppendixGenerator

            # Phase 1: Generate main document
            print("📄 Phase 1: Generating main drilling program...")
            generator = DrillingProgramWordGenerator(self.project)

            if progress:
                progress.setValue(5)

            generator.generate(file_path, progress)

            # Phase 2: Add appendices
            print("📎 Phase 2: Adding appendices and templates...")

            from docx import Document
            doc = Document(file_path)

            appendix_gen = AppendixGenerator(self.project)
            appendix_gen.generate_all_appendices(doc, progress)

            if progress:
                progress.setValue(95)

            # Save final document
            doc.save(file_path)

            if progress:
                progress.setValue(100)

            print(f"✅ Document saved: {file_path}")
            return True

        except ImportError as e:
            print(f"❌ Import Error: {e}")
            print("   Make sure all files are in the same directory")
            # Fallback: generate without appendices
            return self._generate_fallback(file_path, progress)

        except Exception as e:
            print(f"❌ Error: {e}")
            import traceback
            traceback.print_exc()
            return False

    def _generate_fallback(self, file_path: str, progress=None):
        """تولید سند بدون ضمایم (Fallback)"""
        try:
            from word_generator import DrillingProgramWordGenerator
            generator = DrillingProgramWordGenerator(self.project)
            generator.generate(file_path, progress)
            print(f"✅ Document saved (without appendices): {file_path}")
            return True
        except Exception as e:
            print(f"❌ Fallback also failed: {e}")
            return False


# ============================================================================
# SAMPLE PROJECT DATA
# ============================================================================

class SampleProjectGenerator:
    """تولید پروژه نمونه"""

    @staticmethod
    def create_middle_east_development_well():
        """ایجاد پروژه نمونه - چاه توسعه‌ای خاورمیانه"""
        from main import (
            WellProject, CompanyInfo, WellGeneralInfo,
            FormationTop, HazardEntry, CasingDesign,
            CementDesign, BHADesign, MudProgram,
            DirectionalPlan, BOPStack, WellControlData,
            RigSpecification, DrillingParameters, TimeEstimate,
            DrillStringComponent
        )

        project = WellProject()

        # Company Info
        project.company_info = CompanyInfo(
            operator_name="Arabian Gulf Oil Company",
            contractor_name="National Drilling Company",
            field_name="Al-Shaheen",
            well_name="ASH-125H",
            well_number="125",
            pad_name="Pad-12",
            rig_name="NDC Rig 45",
            rig_type="Land Rig",
            country="United Arab Emirates",
            region="Abu Dhabi",
            block_license="ADCO Block 4",
            api_number="UAE-AD-2024-0125",
            spud_date="2024-06-15",
            prepared_by="Ahmad Al-Rashidi, Sr. Drilling Engineer",
            reviewed_by="Mohammed Al-Fahad, Drilling Superintendent",
            approved_by="Khalid Al-Mansouri, Drilling Manager",
            revision="0",
            document_number="DRL-PRG-ASH125-2024-001",
            classification="Confidential"
        )

        # Well Info
        project.well_info = WellGeneralInfo(
            well_type="Development",
            well_profile="Directional J-Type",
            total_depth_md=12500,
            total_depth_tvd=10800,
            water_depth=0,
            air_gap=0,
            ground_elevation=125,
            kb_elevation=28.5,
            rt_elevation=28.5,
            magnetic_declination=2.15,
            grid_convergence=0.85,
            surface_latitude="24° 28' 15.3\" N",
            surface_longitude="54° 22' 48.7\" E",
            target_latitude="24° 28' 42.1\" N",
            target_longitude="54° 23' 05.8\" E",
            coordinate_system="WGS 84",
            target_formation="Arab-D Reservoir",
            target_zone="Zone 2B - Upper Arab-D",
            expected_reservoir_pressure=4850,
            expected_reservoir_temperature=245,
            expected_h2s_concentration=3.5,
            expected_co2_concentration=2.1,
            nace_required=True,
            wellhead_type="Compact",
            xmas_tree_type="Conventional",
            seismic_reference="3D Seismic Survey 2022, Lines AL-45 to AL-52"
        )

        # Formation Tops
        project.formation_tops = [
            FormationTop(
                name="Dibdibba Formation", formation_type="Sand",
                md_top=0, md_bottom=650, tvd_top=0, tvd_bottom=650,
                pore_pressure_top=8.6, pore_pressure_bottom=8.6,
                fracture_gradient_top=14.0, fracture_gradient_bottom=14.5,
                overburden_gradient=16.0,
                temperature_top=95, temperature_bottom=115,
                drillability="Easy", directional_tendency="Neutral",
                remarks="Unconsolidated sand, possible shallow water"
            ),
            FormationTop(
                name="Dammam Formation", formation_type="Limestone",
                md_top=650, md_bottom=2200, tvd_top=650, tvd_bottom=2200,
                pore_pressure_top=8.7, pore_pressure_bottom=8.8,
                fracture_gradient_top=14.5, fracture_gradient_bottom=15.0,
                overburden_gradient=16.5,
                temperature_top=115, temperature_bottom=150,
                drillability="Medium", directional_tendency="Neutral",
                remarks="Possible lost circulation in fractured zones"
            ),
            FormationTop(
                name="Rus Formation", formation_type="Anhydrite",
                md_top=2200, md_bottom=3100, tvd_top=2200, tvd_bottom=3100,
                pore_pressure_top=8.8, pore_pressure_bottom=9.0,
                fracture_gradient_top=15.0, fracture_gradient_bottom=15.5,
                overburden_gradient=17.0,
                temperature_top=150, temperature_bottom=170,
                drillability="Medium", directional_tendency="Neutral",
                remarks="Anhydrite/Marl interbedded"
            ),
            FormationTop(
                name="Umm Er Radhuma", formation_type="Limestone",
                md_top=3100, md_bottom=4800, tvd_top=3100, tvd_bottom=4600,
                pore_pressure_top=8.9, pore_pressure_bottom=9.2,
                fracture_gradient_top=15.5, fracture_gradient_bottom=16.0,
                overburden_gradient=17.5,
                temperature_top=170, temperature_bottom=200,
                drillability="Medium", directional_tendency="Build",
                remarks="KOP in this formation, possible losses in vugs"
            ),
            FormationTop(
                name="Simsima / Aruma", formation_type="Limestone",
                md_top=4800, md_bottom=6500, tvd_top=4600, tvd_bottom=6100,
                pore_pressure_top=9.0, pore_pressure_bottom=9.5,
                fracture_gradient_top=16.0, fracture_gradient_bottom=16.5,
                overburden_gradient=18.0,
                temperature_top=200, temperature_bottom=220,
                drillability="Hard", directional_tendency="Hold",
                remarks="Dense limestone, slow drilling expected"
            ),
            FormationTop(
                name="Wasia Formation", formation_type="Sandstone",
                md_top=6500, md_bottom=8200, tvd_top=6100, tvd_bottom=7500,
                pore_pressure_top=9.2, pore_pressure_bottom=9.8,
                fracture_gradient_top=16.5, fracture_gradient_bottom=17.0,
                overburden_gradient=18.5,
                temperature_top=220, temperature_bottom=235,
                drillability="Medium", directional_tendency="Hold",
                remarks="Sandstone with shale interbeds"
            ),
            FormationTop(
                name="Shuaiba Formation", formation_type="Limestone",
                md_top=8200, md_bottom=9400, tvd_top=7500, tvd_bottom=8400,
                pore_pressure_top=9.5, pore_pressure_bottom=10.0,
                fracture_gradient_top=17.0, fracture_gradient_bottom=17.5,
                overburden_gradient=19.0,
                temperature_top=235, temperature_bottom=240,
                drillability="Hard", directional_tendency="Hold",
                remarks="Dense limestone, mud losses possible"
            ),
            FormationTop(
                name="Kharaib Formation", formation_type="Limestone",
                md_top=9400, md_bottom=10500, tvd_top=8400, tvd_bottom=9200,
                pore_pressure_top=10.0, pore_pressure_bottom=10.3,
                fracture_gradient_top=17.0, fracture_gradient_bottom=17.5,
                overburden_gradient=19.0,
                temperature_top=240, temperature_bottom=242,
                drillability="Hard", directional_tendency="Hold",
                remarks="Transition to reservoir section"
            ),
            FormationTop(
                name="Arab-D Reservoir", formation_type="Dolomite",
                md_top=10500, md_bottom=12500, tvd_top=9200, tvd_bottom=10800,
                pore_pressure_top=10.2, pore_pressure_bottom=10.5,
                fracture_gradient_top=17.0, fracture_gradient_bottom=17.5,
                overburden_gradient=19.5,
                temperature_top=242, temperature_bottom=250,
                drillability="Hard", directional_tendency="Hold",
                remarks="TARGET ZONE - H2S bearing, NACE required"
            ),
        ]

        # Hazards
        project.hazards = [
            HazardEntry(
                hazard_type="Shallow Water Flow",
                md_top=0, md_bottom=500,
                severity="Medium", probability="Possible",
                description="Shallow aquifer in Dibdibba sand",
                mitigation="Diverter installed, controlled ROP",
                contingency="Divert if flow encountered"
            ),
            HazardEntry(
                hazard_type="Lost Circulation",
                md_top=650, md_bottom=2200,
                severity="Medium", probability="Likely",
                description="Fractured/vugular zones in Dammam Fm",
                mitigation="LCM in mud, monitor ECD",
                contingency="LCM pills, cement squeeze if severe"
            ),
            HazardEntry(
                hazard_type="Lost Circulation",
                md_top=3100, md_bottom=4800,
                severity="High", probability="Likely",
                description="Vugular porosity in UER",
                mitigation="Pre-treat with LCM, reduced flow rate",
                contingency="Cement plug, settable pill"
            ),
            HazardEntry(
                hazard_type="Stuck Pipe",
                md_top=4800, md_bottom=8200,
                severity="Medium", probability="Possible",
                description="Differential sticking in permeable zones",
                mitigation="Minimize static time, proper MW, jar in BHA",
                contingency="Spotting pill, free point, backoff"
            ),
            HazardEntry(
                hazard_type="H2S Gas",
                md_top=10500, md_bottom=12500,
                severity="Critical", probability="Almost Certain",
                description=f"H2S concentration ~3.5% in Arab-D",
                mitigation="NACE materials, SCBA, H2S monitors, drills",
                contingency="H2S emergency plan, evacuation procedures"
            ),
            HazardEntry(
                hazard_type="High Temperature",
                md_top=8000, md_bottom=12500,
                severity="Medium", probability="Likely",
                description="BHT up to 250°F",
                mitigation="HT rated tools, thermal stability of mud",
                contingency="Adjust mud chemistry, HT cement"
            ),
        ]

        # Casing Design
        project.casing_design = [
            CasingDesign(
                section_name="Conductor", section_type="Conductor",
                hole_size=36, casing_od=30, casing_id=28.0,
                casing_weight=309.72, casing_grade="X-52",
                casing_connection="Welded",
                setting_depth_md=250, setting_depth_tvd=250,
                shoe_depth_md=250, shoe_depth_tvd=250,
                top_of_cement_md=0, cement_to_surface=True,
                burst_rating=2500, collapse_rating=1000,
                tensile_rating=500000, drift_id=28.0,
                centralizer_type="Rigid", centralizer_spacing=40,
                float_collar_depth=230, float_shoe_type="Guide Shoe",
                remarks="Driven or cemented"
            ),
            CasingDesign(
                section_name="Surface", section_type="Surface",
                hole_size=26, casing_od=20, casing_id=18.73,
                casing_weight=133, casing_grade="K-55",
                casing_connection="BTC",
                setting_depth_md=2500, setting_depth_tvd=2500,
                shoe_depth_md=2500, shoe_depth_tvd=2500,
                top_of_cement_md=0, cement_to_surface=True,
                burst_rating=3060, collapse_rating=1490,
                tensile_rating=1160000, drift_id=18.63,
                min_design_factor_burst=1.10,
                min_design_factor_collapse=1.10,
                min_design_factor_tension=1.60,
                centralizer_type="Bow-Spring", centralizer_spacing=60,
                float_collar_depth=2450,
                remarks="Cement to surface, protect freshwater aquifer"
            ),
            CasingDesign(
                section_name="Intermediate", section_type="Intermediate",
                hole_size=17.5, casing_od=13.375, casing_id=12.415,
                casing_weight=72, casing_grade="N-80",
                casing_connection="BTC",
                setting_depth_md=8200, setting_depth_tvd=7500,
                shoe_depth_md=8200, shoe_depth_tvd=7500,
                top_of_cement_md=5000,
                burst_rating=5020, collapse_rating=2670,
                tensile_rating=1556000, drift_id=12.259,
                min_design_factor_burst=1.10,
                min_design_factor_collapse=1.10,
                min_design_factor_tension=1.60,
                centralizer_type="Bow-Spring", centralizer_spacing=60,
                float_collar_depth=8140,
                remarks="Isolate loss zones and build section"
            ),
            CasingDesign(
                section_name="Production", section_type="Production",
                hole_size=12.25, casing_od=9.625, casing_id=8.535,
                casing_weight=53.5, casing_grade="L-80",
                casing_connection="VAM TOP",
                setting_depth_md=12500, setting_depth_tvd=10800,
                shoe_depth_md=12500, shoe_depth_tvd=10800,
                top_of_cement_md=7000,
                burst_rating=7930, collapse_rating=4750,
                tensile_rating=1243000, drift_id=8.379,
                min_design_factor_burst=1.10,
                min_design_factor_collapse=1.10,
                min_design_factor_tension=1.80,
                centralizer_type="Bow-Spring", centralizer_spacing=40,
                float_collar_depth=12440,
                remarks="NACE L-80, sour service, premium connections"
            ),
        ]

        # Cement Design
        project.cement_design = [
            CementDesign(
                section_name="Surface", casing_od=20, hole_size=26,
                shoe_depth_md=2500, toc_md=0,
                lead_slurry_type="API Class G", lead_slurry_density=13.0,
                lead_slurry_yield=1.48, lead_slurry_volume=350,
                lead_slurry_thickening_time=8,
                lead_slurry_compressive_strength=500,
                tail_slurry_type="API Class G", tail_slurry_density=15.8,
                tail_slurry_yield=1.15, tail_slurry_volume=120,
                tail_slurry_thickening_time=6,
                tail_slurry_compressive_strength=2000,
                spacer_type="Chemical Wash", spacer_density=9.0,
                spacer_volume=30,
                wash_type="Fresh Water", wash_volume=20,
                displacement_volume=450, displacement_rate=8,
                excess_percentage=100, woc_time=12,
                plug_bump_pressure=500, cbl_cbil_required=True,
                cement_additives="Retarder, FL Additive, Dispersant",
                remarks="Cement to surface - verify returns"
            ),
            CementDesign(
                section_name="Intermediate", casing_od=13.375,
                hole_size=17.5, shoe_depth_md=8200, toc_md=5000,
                lead_slurry_type="Class G + Perlite",
                lead_slurry_density=12.5,
                lead_slurry_yield=1.92, lead_slurry_volume=500,
                lead_slurry_thickening_time=10,
                lead_slurry_compressive_strength=300,
                tail_slurry_type="API Class G", tail_slurry_density=15.8,
                tail_slurry_yield=1.15, tail_slurry_volume=180,
                tail_slurry_thickening_time=8,
                tail_slurry_compressive_strength=2500,
                spacer_type="Spacer (Mud Push)", spacer_density=10.5,
                spacer_volume=40,
                wash_type="Chemical Wash", wash_volume=25,
                displacement_volume=520, displacement_rate=6,
                excess_percentage=50, woc_time=18,
                plug_bump_pressure=700, cbl_cbil_required=True,
                cement_additives="Retarder, FL, Anti-Gas Migration",
                remarks="Critical cement job"
            ),
            CementDesign(
                section_name="Production", casing_od=9.625,
                hole_size=12.25, shoe_depth_md=12500, toc_md=7000,
                lead_slurry_type="Class G + Microspheres",
                lead_slurry_density=12.0,
                lead_slurry_yield=2.10, lead_slurry_volume=420,
                lead_slurry_thickening_time=12,
                lead_slurry_compressive_strength=250,
                tail_slurry_type="Class G + Silica Flour",
                tail_slurry_density=16.4,
                tail_slurry_yield=1.08, tail_slurry_volume=150,
                tail_slurry_thickening_time=10,
                tail_slurry_compressive_strength=3000,
                spacer_type="Oil-Based Spacer", spacer_density=11.0,
                spacer_volume=35,
                wash_type="Chemical Wash", wash_volume=20,
                displacement_volume=480, displacement_rate=5,
                excess_percentage=50, woc_time=24,
                plug_bump_pressure=1000, cbl_cbil_required=True,
                cement_additives="Retarder, FL, Gas Block, Silica Flour, H2S Resistant",
                remarks="Critical - reservoir isolation, sour service cement"
            ),
        ]

        # BHA Designs
        project.bha_designs = [
            BHADesign(
                section_name="Surface", bha_number=1,
                hole_size=26, bha_type="Rotary",
                bit_type="Tricone Insert", bit_size=26,
                bit_manufacturer="Smith Bits", bit_model="XR+",
                bit_nozzles="3x22 (1.14 sq.in)",
                mwd_type="Gyro MWD",
                recommended_wob="15-30 klbs", recommended_rpm="80-120",
                recommended_flow_rate="800-1000 GPM",
                recommended_torque="25000 ft-lbs"
            ),
            BHADesign(
                section_name="Intermediate", bha_number=1,
                hole_size=17.5, bha_type="Motor + MWD/LWD",
                bit_type="PDC", bit_size=17.5,
                bit_manufacturer="Baker Hughes", bit_model="Dynamus HCM",
                bit_nozzles="5x16, 2x14 (0.98 sq.in)",
                motor_type="9-5/8\" Positive Displacement Motor",
                motor_od=9.625, motor_bend=1.5,
                mwd_type="OnTrak MWD", lwd_sensors="GR, RES, DEN, NEU",
                recommended_wob="15-35 klbs", recommended_rpm="100-160",
                recommended_flow_rate="700-850 GPM",
                recommended_torque="35000 ft-lbs",
                remarks="Build & Hold section - 3.0°/100ft build rate"
            ),
            BHADesign(
                section_name="Production", bha_number=1,
                hole_size=12.25, bha_type="RSS + MWD/LWD",
                bit_type="PDC", bit_size=12.25,
                bit_manufacturer="Schlumberger", bit_model="Smith SDi616",
                bit_nozzles="4x13, 2x12 (0.68 sq.in)",
                rss_type="PowerDrive X6", rss_model="X6",
                mwd_type="TeleScope MWD",
                lwd_sensors="GR, RES, DEN, NEU, SON, PRESS",
                recommended_wob="10-25 klbs", recommended_rpm="120-180",
                recommended_flow_rate="550-700 GPM",
                recommended_torque="30000 ft-lbs",
                remarks="Reservoir section - H2S zone, minimize formation damage"
            ),
        ]

        # Mud Programs
        project.mud_programs = [
            MudProgram(
                section_name="Surface", hole_size=26,
                depth_from=250, depth_to=2500,
                mud_type="WBM - KCl Polymer",
                mud_weight_in=9.0, mud_weight_out=9.5,
                funnel_viscosity=50, plastic_viscosity=18,
                yield_point=15, gel_strength_10s=6,
                gel_strength_10m=12, gel_strength_30m=18,
                fluid_loss=5.0, hthp_fluid_loss=12.0,
                ph=10.0, chlorides=2000, mbt=15, sand_content=0.3,
                total_volume_required=1500, active_volume=800,
                key_additives="KCl, PHPA, Starch, PAC-R, PAC-L, Soda Ash",
                ecd_at_shoe=9.7, ecd_at_td=9.8,
                remarks="Spud with gel/water then convert to KCl polymer"
            ),
            MudProgram(
                section_name="Intermediate", hole_size=17.5,
                depth_from=2500, depth_to=8200,
                mud_type="WBM - KCl Polymer",
                mud_weight_in=9.5, mud_weight_out=12.0,
                funnel_viscosity=55, plastic_viscosity=22,
                yield_point=18, gel_strength_10s=8,
                gel_strength_10m=14, gel_strength_30m=20,
                fluid_loss=4.0, hthp_fluid_loss=8.0,
                ph=10.5, chlorides=5000, mbt=20, sand_content=0.2,
                total_volume_required=2500, active_volume=1200,
                key_additives="KCl, PHPA, Starch, PAC, Barite, LCM, Lubricant",
                ecd_at_shoe=12.3, ecd_at_td=12.5,
                remarks="Increase MW through loss zones gradually"
            ),
            MudProgram(
                section_name="Production", hole_size=12.25,
                depth_from=8200, depth_to=12500,
                mud_type="OBM - Mineral Oil Based",
                mud_weight_in=10.5, mud_weight_out=12.8,
                funnel_viscosity=55, plastic_viscosity=20,
                yield_point=12, gel_strength_10s=6,
                gel_strength_10m=10, gel_strength_30m=14,
                fluid_loss=0, hthp_fluid_loss=4.0,
                ph=0, chlorides=0, mbt=0, sand_content=0.1,
                oil_water_ratio="80/20",
                electrical_stability=800,
                total_volume_required=3000, active_volume=1500,
                key_additives="Mineral Oil, CaCl2, Organophilic Clay, "
                              "Barite, Lime, Primary Emulsifier, "
                              "Wetting Agent, Fluid Loss Additive",
                ecd_at_shoe=13.0, ecd_at_td=13.2,
                remarks="OBM for reservoir - minimize formation damage, "
                        "H2S compatible"
            ),
        ]

        # Directional Plan
        project.directional_plan = DirectionalPlan(
            survey_tool="MWD (Magnetic) + Gyro at shoes",
            survey_frequency=90,
            kickoff_point_md=3500, kickoff_point_tvd=3500,
            build_rate=3.0, turn_rate=0,
            hold_inclination=35, hold_azimuth=45,
            target_inclination=35, target_azimuth=45,
            max_dls=5.0,
            horizontal_displacement=4200,
            vertical_section=4200,
        )

        # BOP
        project.bop_stack = BOPStack(
            bop_type="Triple Ram",
            working_pressure=10000,
            bore_size=18.75,
            manufacturer="Cameron",
            model="Type U",
            annular_preventer_size=18.75,
            annular_preventer_wp=10000,
            pipe_ram_size='5", 3-1/2"',
            blind_ram=True, shear_ram=True,
            variable_bore_ram=True,
            kill_line_size=3.0, choke_line_size=3.0,
            choke_manifold_wp=10000,
            accumulator_capacity=1600,
            accumulator_precharge=1000,
            diverter_size=30, diverter_line_size=12,
            function_test_frequency="Weekly",
            pressure_test_frequency="Per Section",
            bop_test_pressure_low=250,
            bop_test_pressure_high=7000,
        )

        # Well Control
        project.well_control = WellControlData(
            maasp_surface=2500,
            kill_method="Wait & Weight (Engineer's Method)",
            kick_tolerance=50,
            slow_pump_rate_1=30,
            slow_pump_pressure_1=850,
            slow_pump_rate_2=20,
            slow_pump_pressure_2=620,
            pit_gain_action_level=5,
            h2s_action_levels="10ppm: Alert, 20ppm: Alarm/SCBA, "
                              "50ppm: Evacuate, 100ppm: Emergency",
            emergency_contacts="Company Man: +971-XX-XXXXXXX\n"
                               "Base Office: +971-XX-XXXXXXX\n"
                               "Hospital: Al-Ain Hospital +971-XX-XXXXXXX\n"
                               "Fire: 997\nPolice: 999\nAmbulance: 998"
        )

        # Rig Spec
        project.rig_spec = RigSpecification(
            rig_name="NDC Rig 45",
            rig_type="Land Rig",
            rig_contractor="National Drilling Company",
            max_hook_load=1000000,
            drawworks_power=2000,
            top_drive=True,
            top_drive_model="NOV TDS-11SA",
            top_drive_torque=37500,
            max_rotary_speed=250,
            derrick_height=147,
            rotary_table_size=37.5,
            mud_pump_1_type="NOV 14-P-220",
            mud_pump_1_hp=1600,
            mud_pump_1_liner=6.5,
            mud_pump_1_max_pressure=5000,
            mud_pump_1_max_flow=850,
            mud_pump_2_type="NOV 14-P-220",
            mud_pump_2_hp=1600,
            mud_pump_2_liner=6.5,
            mud_pump_2_max_pressure=5000,
            mud_pump_2_max_flow=850,
            mud_pump_3_type="NOV 14-P-220 (Standby)",
            mud_pump_3_hp=1600,
            pit_volume_total=1800,
            pit_volume_active=1200,
            shale_shaker_count=4,
            degasser_type="Centrifugal Degasser",
            centrifuge="2 x Decanting Centrifuge 518",
            generators="4 x CAT 3512B, 1500 kW each",
            total_power=6000,
            crane_capacity=50,
            accommodation=120,
        )

        # Time Estimates
        project.time_estimates = [
            TimeEstimate(section_name="Pre-Spud",
                         operation="Rig Move & Setup",
                         total_section_days=3.0, cumulative_days=3.0),
            TimeEstimate(section_name="Conductor",
                         operation="Drill & Set 30\" Conductor",
                         depth_from=0, depth_to=250,
                         total_section_days=2.0, cumulative_days=5.0),
            TimeEstimate(section_name="Surface",
                         operation="Drill 26\" & Set 20\" Casing",
                         depth_from=250, depth_to=2500,
                         rop_average=40,
                         total_section_days=6.0, cumulative_days=11.0),
            TimeEstimate(section_name="Intermediate",
                         operation="Drill 17-1/2\" & Set 13-3/8\"",
                         depth_from=2500, depth_to=8200,
                         rop_average=20,
                         total_section_days=18.0, cumulative_days=29.0),
            TimeEstimate(section_name="Production",
                         operation="Drill 12-1/4\" & Set 9-5/8\"",
                         depth_from=8200, depth_to=12500,
                         rop_average=12,
                         total_section_days=22.0, cumulative_days=51.0),
            TimeEstimate(section_name="Completion",
                         operation="Completion & Testing",
                         total_section_days=7.0, cumulative_days=58.0),
        ]

        # Drilling Parameters
        project.drilling_parameters = [
            DrillingParameters(
                section_name="Surface", hole_size=26,
                depth_from=250, depth_to=2500,
                wob_min=15, wob_max=30,
                rpm_min=80, rpm_max=120,
                flow_rate_min=800, flow_rate_max=1000,
                torque_max=25000, rop_average=40,
                spp_max=2500, overpull_limit=50,
                max_ecd=14.5
            ),
            DrillingParameters(
                section_name="Intermediate", hole_size=17.5,
                depth_from=2500, depth_to=8200,
                wob_min=15, wob_max=35,
                rpm_min=100, rpm_max=160,
                flow_rate_min=700, flow_rate_max=850,
                torque_max=35000, rop_average=20,
                spp_max=3500, overpull_limit=50,
                max_ecd=16.0
            ),
            DrillingParameters(
                section_name="Production", hole_size=12.25,
                depth_from=8200, depth_to=12500,
                wob_min=10, wob_max=25,
                rpm_min=120, rpm_max=180,
                flow_rate_min=550, flow_rate_max=700,
                torque_max=30000, rop_average=12,
                spp_max=4000, overpull_limit=50,
                max_ecd=17.0
            ),
        ]

        return project


# ============================================================================
# CONFIGURATION MANAGER
# ============================================================================

class ConfigManager:
    """مدیریت تنظیمات"""

    DEFAULT_CONFIG = {
        "app": {
            "name": "Drilling Program Generator Pro",
            "version": "3.0",
            "language": "en",
            "theme": "dark",
            "auto_save": True,
            "auto_save_interval": 300,
            "recent_files_max": 10,
        },
        "document": {
            "paper_size": "A4",
            "orientation": "Portrait",
            "font_family": "Calibri",
            "font_size": 10,
            "header_font_size": 16,
            "include_appendices": True,
            "include_kill_sheet": True,
            "include_rig_poster": True,
            "include_ddr_template": True,
            "include_trip_sheet": True,
            "include_tally_sheet": True,
            "include_bit_record": True,
            "include_survey_sheet": True,
            "include_cement_report": True,
            "include_cost_estimate": True,
        },
        "units": {
            "depth": "ft",
            "pressure": "psi",
            "mud_weight": "ppg",
            "temperature": "F",
            "flow_rate": "GPM",
            "volume": "bbl",
            "weight": "lbs",
            "torque": "ft-lbs",
            "length": "ft",
        },
        "design_factors": {
            "burst_min": 1.10,
            "collapse_min": 1.10,
            "tension_min": 1.60,
            "triaxial_min": 1.25,
        },
        "standards": {
            "casing_design": "API 5C3 / ISO 10400",
            "cement": "API RP 10B / ISO 10426",
            "mud": "API RP 13B / 13D",
            "well_control": "API RP 53 / 59",
            "bop_testing": "API RP 53",
            "directional": "ISCWSA / SPE",
        },
        "paths": {
            "projects_dir": "projects",
            "exports_dir": "projects/exports",
            "templates_dir": "projects/templates",
            "backup_dir": "projects/backup",
            "logs_dir": "logs",
        }
    }

    @staticmethod
    def load_config(config_path: str = "config/settings.json") -> dict:
        """بارگذاری تنظیمات"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            return {**ConfigManager.DEFAULT_CONFIG, **config}
        except (FileNotFoundError, json.JSONDecodeError):
            return ConfigManager.DEFAULT_CONFIG

    @staticmethod
    def save_config(config: dict,
                    config_path: str = "config/settings.json"):
        """ذخیره تنظیمات"""
        Path(config_path).parent.mkdir(parents=True, exist_ok=True)
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)

    @staticmethod
    def initialize():
        """راه‌اندازی اولیه"""
        ProjectStructure.create_project_structure()
        config = ConfigManager.DEFAULT_CONFIG
        ConfigManager.save_config(config)
        print("✅ Configuration initialized")


# ============================================================================
# COMMAND LINE INTERFACE
# ============================================================================

class CLI:
    """رابط خط فرمان"""

    @staticmethod
    def print_banner():
        """نمایش بنر"""
        banner = """
╔══════════════════════════════════════════════════════════════╗
║                                                              ║
║    🛢️  DRILLING PROGRAM GENERATOR PRO v3.0                   ║
║                                                              ║
║    Professional Drilling Program & Procedure Generator        ║
║    Based on API, IADC, Shell DEP, Saudi Aramco Standards     ║
║                                                              ║
║    Features:                                                  ║
║    ✅ Complete Drilling Program Document                      ║
║    ✅ 25+ Detailed Operating Procedures                      ║
║    ✅ Engineering Calculations (Hydraulics, T&D, Cement)     ║
║    ✅ Kill Sheet, Trip Sheet, Tally Sheet                    ║
║    ✅ Daily Report Template (IADC Format)                    ║
║    ✅ Rig Floor Poster                                       ║
║    ✅ Bit Record, Survey Sheet, Cement Report                ║
║    ✅ Professional Word Document Output                      ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝
        """
        print(banner)

    @staticmethod
    def run():
        """اجرای CLI"""
        CLI.print_banner()

        if len(sys.argv) > 1:
            command = sys.argv[1].lower()

            if command == "--install":
                print("\n📦 Installing dependencies...")
                DependencyManager.install_dependencies(
                    include_optional="--all" in sys.argv
                )

            elif command == "--check":
                DependencyManager.print_status()

            elif command == "--init":
                ConfigManager.initialize()

            elif command == "--sample":
                CLI.generate_sample()

            elif command == "--gui":
                CLI.launch_gui()

            elif command == "--help":
                CLI.print_help()

            else:
                print(f"Unknown command: {command}")
                CLI.print_help()
        else:
            # Default: Launch GUI
            CLI.launch_gui()

    @staticmethod
    def launch_gui():
        """راه‌اندازی رابط گرافیکی"""
        print("\n🚀 Launching GUI...")

        # Check dependencies first
        if not DependencyManager.print_status():
            print("\n❌ Please install required dependencies first:")
            print("   python launcher.py --install")
            return

        try:
            from PySide6.QtWidgets import QApplication
            from main import DrillingProgramMainWindow, DARK_STYLE

            app = QApplication(sys.argv)
            app.setStyle('Fusion')
            app.setStyleSheet(DARK_STYLE)
            app.setApplicationName("Drilling Program Generator Pro")
            app.setApplicationVersion("3.0")

            window = DrillingProgramMainWindow()
            window.show()

            print("✅ Application started successfully")
            sys.exit(app.exec())

        except ImportError as e:
            print(f"\n❌ Failed to launch GUI: {e}")
            print("   Install dependencies: python launcher.py --install")

    @staticmethod
    def generate_sample():
        """تولید سند نمونه بدون GUI"""
        print("\n📄 Generating sample drilling program...")

        if not DependencyManager.check_dependencies()['all_required_met']:
            print("❌ Required dependencies not met. "
                  "Run: python launcher.py --install")
            return

        try:
            # Create sample project
            project = SampleProjectGenerator.create_middle_east_development_well()

            # Generate document
            output_file = (
                f"projects/exports/"
                f"Drilling_Program_{project.company_info.well_name}_"
                f"Rev{project.company_info.revision}_"
                f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            )

            Path("projects/exports").mkdir(parents=True, exist_ok=True)

            generator = EnhancedWordGenerator(project)
            success = generator.generate_complete_document(output_file)

            if success:
                print(f"\n✅ Sample document generated successfully!")
                print(f"   📁 File: {output_file}")
                print(f"\n   The document includes:")
                print(f"   • Cover Page")
                print(f"   • Table of Contents")
                print(f"   • Executive Summary")
                print(f"   • Well Information")
                print(f"   • Rig Specifications")
                print(f"   • Formation Prognosis (9 formations)")
                print(f"   • Hazard Analysis (6 hazards)")
                print(f"   • Casing Design (4 strings)")
                print(f"   • Mud Program (3 sections)")
                print(f"   • BHA Design (3 BHAs)")
                print(f"   • Cement Program (3 jobs)")
                print(f"   • Directional Plan")
                print(f"   • BOP & Well Control")
                print(f"   • Time Estimate")
                print(f"   • 25+ Detailed Procedures")
                print(f"   • Kill Sheet")
                print(f"   • Wellbore Schematic")
                print(f"   • Trip Sheet Template")
                print(f"   • Bit Record Template")
                print(f"   • Casing Tally Sheets")
                print(f"   • Survey Sheet Template")
                print(f"   • Cement Job Report Template")
                print(f"   • Cost Estimate Template")
                print(f"   • Daily Drilling Report Template")

                # Open file
                if sys.platform == 'win32':
                    os.startfile(output_file)
                elif sys.platform == 'darwin':
                    os.system(f'open "{output_file}"')
                else:
                    os.system(f'xdg-open "{output_file}"')
            else:
                print("❌ Failed to generate document")

        except Exception as e:
            print(f"❌ Error: {e}")
            import traceback
            traceback.print_exc()

    @staticmethod
    def print_help():
        """نمایش راهنما"""
        help_text = """
USAGE:
    python launcher.py [command]

COMMANDS:
    (none)          Launch GUI application
    --gui           Launch GUI application
    --install       Install required dependencies
    --install --all Install all dependencies (including optional)
    --check         Check dependency status
    --init          Initialize project structure and config
    --sample        Generate sample drilling program document
    --help          Show this help message

EXAMPLES:
    python launcher.py                  # Launch GUI
    python launcher.py --install        # Install dependencies
    python launcher.py --sample         # Generate sample document
    python launcher.py --check          # Check dependencies

FILES:
    main.py                    - Part 1: GUI & Data Models
    engineering_calculations.py - Part 2: Calculations & Procedures
    word_generator.py          - Part 3: Word Document Generator
    advanced_modules.py        - Part 4: Templates & Appendices
    launcher.py                - Part 5: Setup & Integration

REQUIREMENTS:
    Python 3.9+
    PySide6
    python-docx
        """
        print(help_text)


# ============================================================================
# QUICK START SCRIPT
# ============================================================================

def quick_setup():
    """راه‌اندازی سریع"""
    print("\n🔧 QUICK SETUP")
    print("=" * 50)

    # Step 1: Create structure
    print("\n1️⃣  Creating project structure...")
    ProjectStructure.create_project_structure()

    # Step 2: Initialize config
    print("2️⃣  Initializing configuration...")
    ConfigManager.initialize()

    # Step 3: Check dependencies
    print("3️⃣  Checking dependencies...")
    results = DependencyManager.check_dependencies()

    if not results['all_required_met']:
        print("\n4️⃣  Installing missing dependencies...")
        DependencyManager.install_dependencies()
    else:
        print("4️⃣  All dependencies are met! ✅")

    print("\n" + "=" * 50)
    print("✅ Setup complete!")
    print("\nTo start the application:")
    print("  python launcher.py          (GUI mode)")
    print("  python launcher.py --sample (Generate sample)")
    print("=" * 50)


# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--setup":
        quick_setup()
    else:
        CLI.run()