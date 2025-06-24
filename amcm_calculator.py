import streamlit as st
import pandas as pd
import numpy as np
import math
import os

# AMCM Model Data
class AMCMModel:
    def __init__(self, mission_type, number, spec, sd):
        self.mission_type = mission_type
        self.number = number
        self.spec = spec
        self.sd = sd

# AMCM Database - only space-related mission types
AMCM_MODELS = [
    AMCMModel("Spacecraft - Planetary Lander", 3, 2.46, 0.0664),
    AMCMModel("Spacecraft - Planetary", 11, 2.39, 0.0733),
    AMCMModel("Spacecraft - Manned Reentry", 6, 2.27, 0.0815),
    AMCMModel("Spacecraft - Communication", 9, 2.22, 0.0454),
    AMCMModel("Spacecraft - Weather", 6, 2.18, 0.1381),
    AMCMModel("Spacecraft - Physics & Astronomy", 11, 2.17, 0.1109),
    AMCMModel("Spacecraft - Earth Observation", 3, 2.16, 0.0817),
    AMCMModel("Spacecraft - Lunar Rover", 1, 2.14, 0.0),
    AMCMModel("Spacecraft - Manned Habitat", 4, 2.13, 0.0759),
    AMCMModel("Space Transport - Unmanned Reentry", 4, 1.91, 0.0987),
    AMCMModel("Space Transport - Launch Vehicle Stage", 3, 2.01, 0.1128),
    AMCMModel("Space Transport - Upper Stage", 5, 2.07, 0.1147),
    AMCMModel("Space Transport - Liquid Rocket Engine - Lox/Lh", 3, 2.19, 0.1195),
    AMCMModel("Space Transport - Liquid Rocket Engine - Lox/RP-1", 2, 1.84, 0.0189),
    AMCMModel("Space Transport - Payload Fairing", 3, 1.15, 0.0242),
    AMCMModel("Space Transport - Centaur Fairing", 2, 1.6, 0.0692),
]

# AMCM Constants from the Paper
A = 0.000504839
B = 0.594183076
C = 0.653947922
D = 76.99939424
E = 1.68051e-52
F = -0.355322218
G = 1.554982942

def load_inflation_data():
    """Load inflation data from Excel file"""
    try:
        if os.path.exists('Inflation Table.xlsx'):
            df = pd.read_excel('Inflation Table.xlsx', header=None)
            # Get year row and index row (rows 5 and 7 as per the existing code)
            year_row = df.iloc[5].tolist()[1:]  # Skip first column
            index_row = df.iloc[7].tolist()[1:]  # Skip first column
            
            # Remove non-numeric years (e.g., 'TQ')
            year_index_pairs = [(y, idx) for y, idx in zip(year_row, index_row) 
                              if isinstance(y, (int, float)) and not pd.isna(y)]
            
            years = [int(y) for y, _ in year_index_pairs]
            indices = [float(idx) for _, idx in year_index_pairs]
            
            return dict(zip(years, indices))
        else:
            # Fallback inflation data if file not found
            st.warning("Inflation Table.xlsx not found. Using fallback data.")
            return {
                1999: 1,
                2000: 1.040,
                2010: 1.384,
                2020: 1.666,
                2024: 1.857,
                2025: 1.906
            }
    except Exception as e:
        st.error(f"Error loading inflation data: {e}")
        return {2024: 1.857, 2025: 1.906}  # Minimal fallback

def calculate_amcm_cost(quantity, weight, mission_type_index, ioc_year, block_number, difficulty_index):
    """
    Calculate AMCM cost using the formula from the paper.
    
    Formula: a * Q^b * W^c * d^S * e^(1/(IOC-1900)) * B^f * g^D
    """
    try:
        # Get mission type specifications
        model = AMCM_MODELS[mission_type_index]
        S = model.spec
        
        # Difficulty adjustment (-2 to +2, where 0 is average)
        difficulty_factor = difficulty_index - 2
        
        # Calculate cost using the original formula
        cost = (A * 
                math.pow(quantity, B) * 
                math.pow(weight, C) * 
                math.pow(D, S) *  # D is the constant from above
                math.pow(E, 1/(ioc_year - 1900)) * 
                math.pow(block_number, F) * 
                math.pow(G, difficulty_factor) # difficulty_factor is -2 to +2
                )
        
        return cost
    except Exception as e:
        st.error(f"Error in cost calculation: {e}")
        return 0

def main():
    st.set_page_config(
        page_title="AMCM Calculator with Inflation Adjustment",
        page_icon="🚀",
        layout="wide"
    )
    
    st.title("🚀 Advanced Missions Cost Model (AMCM) Calculator")
    st.markdown("### With Inflation Adjustment Using NASA New Start Inflation Index")
    
    st.markdown("""
    This calculator provides a useful method for quick turnaround, rough-order-of-magnitude estimating. 
    The model can be used for estimating the development and production cost of spacecrafts.
    """)
    
    # Load inflation data
    inflation_data = load_inflation_data()
    available_years = sorted(inflation_data.keys())
    
    # Create two columns for input and results
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.header("📊 Input Parameters")
        # Mass unit selection
        mass_unit = st.radio(
            "Mass Unit",
            ["lbs", "kg"],
            index=0,
            help="Choose the unit for entering dry mass. Conversion is automatic."
        )
        # Weight input (accepts either lbs or kg)
        if mass_unit == "lbs":
            weight = st.number_input(
                "Dry Weight (lbs)",
                min_value=0.1,
                value=1000.0,
                help="The total empty weight of the system in pounds, not including fuel, payload, crew, or passengers"
            )
            weight_lbs = weight
        else:
            weight = st.number_input(
                "Dry Weight (kg)",
                min_value=0.05,
                value=453.6,
                help="The total empty weight of the system in kilograms, not including fuel, payload, crew, or passengers"
            )
            weight_lbs = weight * 2.20462
        
        # Quantity input
        quantity = st.number_input(
            "Quantity",
            min_value=1,
            value=1,
            help="The total number of units to be produced"
        )
        
        # Mission type selection (space-related only)
        mission_types = [model.mission_type for model in AMCM_MODELS]
        mission_type_index = st.selectbox(
            "Mission Type",
            range(len(mission_types)),
            format_func=lambda x: mission_types[x],
            help="Select the mission type that best describes the system you wish to estimate"
        )
        
        # IOC Year
        ioc_year = st.number_input(
            "IOC Year (Initial Operating Capability)",
            min_value=1901,
            max_value=2100,
            value=2025,
            help="The year in which the spacecraft or vehicle is first launched"
        )
        
        # Block Number
        block_number = st.number_input(
            "Block Number",
            min_value=1,
            value=1,
            help="Level of design inheritance (1 = new design, 2+ = modification of existing design)"
        )
        
        # Difficulty
        difficulty_options = ["Very Low", "Low", "Average", "High", "Very High"]
        difficulty_index = st.selectbox(
            "Difficulty",
            range(len(difficulty_options)),
            index=2,  # Default to "Average"
            format_func=lambda x: difficulty_options[x],
            help="Level of programmatic and technical difficulty relative to similar past systems"
        )
    
    with col2:
        st.header("💰 Cost Results")
        
        # Calculate base cost (1999 dollars)
        base_cost = calculate_amcm_cost(quantity, weight_lbs, mission_type_index, ioc_year, block_number, difficulty_index)
        
        # Display base cost
        st.metric("Base Cost (1999 $)", f"${base_cost:.2f} million")
        
        # Inflation adjustment section
        st.subheader("🔄 Inflation Adjustment")
        
        col2a, col2b = st.columns(2)
        
        with col2a:
            base_year = st.selectbox(
                "Base Year (Cost Year)",
                available_years,
                index=available_years.index(1999) if 1999 in available_years else 0,
                help="Year the base cost is calculated for"
            )
        
        with col2b:
            target_year = st.selectbox(
                "Target Year",
                available_years,
                index=available_years.index(2025) if 2025 in available_years else -1,
                help="Year to escalate costs to"
            )
        
        # Calculate inflation factor
        base_index = inflation_data.get(base_year, 1)
        target_index = inflation_data.get(target_year, 1)
        inflation_factor = target_index / base_index if base_index else 1
        
        st.info(f"Inflation factor from {base_year} to {target_year}: {inflation_factor:.3f}")
        
        # Calculate adjusted cost
        adjusted_cost = base_cost * inflation_factor
        
        # Display adjusted cost
        st.metric(
            f"Adjusted Cost ({target_year} $)", 
            f"${adjusted_cost:.2f} million",
            delta=f"${adjusted_cost - base_cost:.2f} million"
        )
        
        # Cost breakdown
        st.subheader("📈 Cost Breakdown")
        breakdown_data = {
            "Cost Component": ["Base Cost (1999 $)", f"Inflation Adjustment ({base_year} → {target_year})", f"Total Cost ({target_year} $)"],
            "Amount ($ millions)": [f"{base_cost:.2f}", f"{adjusted_cost - base_cost:.2f}", f"{adjusted_cost:.2f}"]
        }
        st.table(pd.DataFrame(breakdown_data))
    
    # Additional information section
    st.markdown("---")
    st.header("📋 Model Information")
    
    col3, col4 = st.columns(2)
    
    with col3:
        st.subheader("Selected Mission Type Details")
        selected_model = AMCM_MODELS[mission_type_index]
        model_info = {
            "Mission Type": selected_model.mission_type,
            "Number of Data Points": selected_model.number,
            "Specification Factor": f"{selected_model.spec:.3f}",
            "Standard Deviation": f"{selected_model.sd:.4f}" if selected_model.sd > 0 else "N/A"
        }
        for key, value in model_info.items():
            st.write(f"**{key}:** {value}")
    
    with col4:
        st.subheader("Calculation Parameters")
        calc_params = {
            "Quantity": quantity,
            "Weight (lbs)": f"{weight_lbs:,.1f}",
            "Weight (kg)": f"{weight_lbs/2.20462:,.1f}",
            "IOC Year": ioc_year,
            "Block Number": block_number,
            "Difficulty": difficulty_options[difficulty_index],
            "Difficulty Index": difficulty_index - 2
        }
        for key, value in calc_params.items():
            st.write(f"**{key}:** {value}")
    
    # Definitions section
    st.markdown("---")
    with st.expander("📖 Definitions and Help"):
        st.markdown("""
        ### Parameter Definitions
        
        **Quantity**: The total number of units to be produced. This includes prototypes, test articles, operational units, and spares.
        
        **Dry Weight**: The total empty weight of the system in pounds, not including fuel, payload, crew, or passengers.
        
        **Mission Type**: Classifies the type of system by the operating environment and the type of mission to be performed.
        
        **IOC Year**: The year of Initial Operating Capability. For space systems, this is the year in which the spacecraft or vehicle is first launched.
        
        **Block Number**: Represents the level of design inheritance in the system. If the system is a new design, then the block number is 1. If the estimate represents a modification to an existing design, then a block number of 2 or more may be used.
        
        **Difficulty**: The level of programmatic and technical difficulty anticipated for the new system, assessed relative to other similar systems developed in the past.
        
        ### About the Model
        
        The Advanced Missions Cost Model (AMCM) is based on historical cost data and uses parametric relationships to estimate development and production costs. The model uses the following formula:
        
        `Cost = a × Q^b × W^c × d^S × e^(1/(IOC-1900)) × B^f × g^D`
        
        Where:
        - Q = Quantity
        - W = Weight
        - S = Mission type specification factor
        - IOC = Initial Operating Capability year
        - B = Block number
        - D = Difficulty factor
        - Constants: a, b, c, d, e, f, g
        
        ### Inflation Adjustment
        
        The inflation adjustment uses the NASA New Start Inflation Index to convert costs between different years. This provides more accurate cost projections for planning and budgeting purposes.
        """)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: gray;'>
    <p>Advanced Missions Cost Model (AMCM) Calculator</p>
    <p>Based on NASA cost estimation methodologies with inflation adjustment capability</p>
    <p>⚠️ This tool provides rough-order-of-magnitude estimates for planning purposes only</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()