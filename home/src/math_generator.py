import math
import matplotlib.pyplot as plt
import os

def format_scientific(value):
    if value == 0:
        return "0"
    exp = int(math.floor(math.log10(value)))
    coef = value / (10**exp)
    if exp >= -2 and exp <= 2:
        return f"{value:.3g}"
    return f"{coef:.2f} \\times 10^{{{exp}}}"

def generate_math_images(source_type, activity, col_t, col_h, pb_t, pb_h, soil_t, soil_h, dist_top, dist_left, pipe_length, pipe_width, output_dir):
    # Set gamma constant based on source type
    if "Ir-192" in source_type or "Ir" in source_type:
        gamma = 4800
    elif "Co-60" in source_type or "Co" in source_type:
        gamma = 13000
    else:
        # Default to Se-75
        gamma = 2030

    # Ensure output dir exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    plt.rc('text', usetex=False)
    plt.rc('mathtext', fontset='cm')

    # Data structure for the 8 equations
    # each tuple: (limit, limit_str, exp_terms_str, exp_val)
    # limit: 10 (관리), 1 (감시)
    
    val_col = col_t / col_h if col_h else 0
    val_pb = pb_t / pb_h if pb_h else 0
    val_soil = soil_t / soil_h if soil_h else 0

    eq_configs = [
        # 방사선관리구역 (1~4)
        {
            "filename": "eq1.jpg",
            "limit": 10,
            "exp_terms": f"\\frac{{{col_t}}}{{{col_h}}} + \\frac{{{pb_t}}}{{{pb_h}}} + \\frac{{{soil_t}}}{{{soil_h}}}",
            "exp_val": val_col + val_pb + val_soil
        },
        {
            "filename": "eq2.jpg",
            "limit": 10,
            "exp_terms": f"\\frac{{{col_t}}}{{{col_h}}} + \\frac{{{pb_t}}}{{{pb_h}}}",
            "exp_val": val_col + val_pb
        },
        {
            "filename": "eq3.jpg",
            "limit": 10,
            "exp_terms": f"\\frac{{{pb_t}}}{{{pb_h}}} + \\frac{{{soil_t}}}{{{soil_h}}}",
            "exp_val": val_pb + val_soil
        },
        {
            "filename": "eq4.jpg",
            "limit": 10,
            "exp_terms": f"\\frac{{{pb_t}}}{{{pb_h}}} + \\frac{{{col_t}}}{{{col_h}}}",
            "exp_val": val_pb + val_col
        },
        # 방사선감시구역 (5~8)
        {
            "filename": "eq5.jpg",
            "limit": 1,
            "exp_terms": f"\\frac{{{col_t}}}{{{col_h}}} + \\frac{{{pb_t}}}{{{pb_h}}} + \\frac{{{soil_t}}}{{{soil_h}}}",
            "exp_val": val_col + val_pb + val_soil
        },
        {
            "filename": "eq6.jpg",
            "limit": 1,
            "exp_terms": f"\\frac{{{col_t}}}{{{col_h}}} + \\frac{{{pb_t}}}{{{pb_h}}}",
            "exp_val": val_col + val_pb
        },
        {
            "filename": "eq7.jpg",
            "limit": 1,
            "exp_terms": f"\\frac{{{pb_t}}}{{{pb_h}}} + \\frac{{{soil_t}}}{{{soil_h}}}",
            "exp_val": val_pb + val_soil
        },
        {
            "filename": "eq8.jpg",
            "limit": 1,
            "exp_terms": f"\\frac{{{pb_t}}}{{{pb_h}}} + \\frac{{{col_t}}}{{{col_h}}}",
            "exp_val": val_pb + val_col
        }
    ]

    generated_paths = []

    for cfg in eq_configs:
        limit = cfg["limit"]
        exp_terms = cfg["exp_terms"]
        exp_val = cfg["exp_val"]
        
        # Calculate result
        numerator = gamma * activity * math.exp(-0.693 * exp_val)
        result = math.sqrt(numerator / limit)
        result_str = format_scientific(result)

        # Build LaTeX string
        limit_str = "10\\mu Sv/hr" if limit == 10 else "1\\mu Sv/hr"
        eq = r'$\sqrt{\frac{' + f"{gamma}\\mu Sv m^2/Ci\\cdot hr \\times {activity}Ci \\times e^{{-0.693({exp_terms})}}" + r'}{' + limit_str + r'}} = ' + result_str + r'm$'

        fig = plt.figure(figsize=(10, 2))
        # Remove whitespace around the figure
        plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
        fig.text(0.5, 0.5, eq, fontsize=22, ha='center', va='center', color='black')
        plt.axis('off')
        
        out_path = os.path.join(output_dir, cfg["filename"])
        # bbox_inches='tight' tightly crops the whitespace
        plt.savefig(out_path, bbox_inches='tight', pad_inches=0.1, dpi=200)
        plt.close(fig)
        
        generated_paths.append(out_path)

    # Calculate dose rates (평가결과)
    d_ab = (pipe_width / 2.0 + dist_top) / 1000.0  # in meters
    d_cd = (pipe_length / 2.0 + dist_left) / 1000.0 # in meters

    numerator_ab = gamma * activity * math.exp(-0.693 * (val_col + val_pb + val_soil))
    numerator_cd = gamma * activity * math.exp(-0.693 * (val_col + val_pb))

    dose_ab = numerator_ab / (d_ab ** 2) if d_ab > 0 else 0
    dose_cd = numerator_cd / (d_cd ** 2) if d_cd > 0 else 0

    dose_str_ab = format_scientific(dose_ab)
    dose_str_cd = format_scientific(dose_cd)

    eval_configs = [
        {"filename": "eval1.jpg", "dose_str": dose_str_ab},
        {"filename": "eval2.jpg", "dose_str": dose_str_cd},
    ]

    for cfg in eval_configs:
        eq = r'$' + cfg["dose_str"] + r'\mu Sv / hr$'
        fig = plt.figure(figsize=(3, 1))
        plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
        fig.text(0.5, 0.5, eq, fontsize=22, ha='center', va='center', color='black')
        plt.axis('off')
        
        out_path = os.path.join(output_dir, cfg["filename"])
        plt.savefig(out_path, bbox_inches='tight', pad_inches=0.1, dpi=200)
        plt.close(fig)
        
        generated_paths.append(out_path)

    # Calculate full equations
    # We want: 2030 \mu Sv \cdot m^2 / Ci \cdot hr \times 60 Ci \times \frac{1}{(2.5)^2} \times e^{-0.693(\frac{12}{1} + \frac{11}{0.8} + \frac{983}{45})} = 9.24 \times 10^{-11} \mu Sv / hr
    def format_distance(d):
        if d == int(d):
            return str(int(d))
        return str(d)

    d_ab_str = format_distance(d_ab)
    d_cd_str = format_distance(d_cd)

    exp_terms_ab_str = f"\\frac{{{pb_t:g}}}{{{pb_h:g}}} + \\frac{{{col_t:g}}}{{{col_h:g}}} + \\frac{{{soil_t:g}}}{{{soil_h:g}}}"
    exp_terms_cd_str = f"\\frac{{{col_t:g}}}{{{col_h:g}}} + \\frac{{{pb_t:g}}}{{{pb_h:g}}}"

    eval_eq_configs = [
        {"filename": "eval_eq1.jpg", "dist": d_ab_str, "exp_terms": exp_terms_ab_str, "dose_str": dose_str_ab},
        {"filename": "eval_eq2.jpg", "dist": d_cd_str, "exp_terms": exp_terms_cd_str, "dose_str": dose_str_cd},
    ]

    for cfg in eval_eq_configs:
        eq = r'$' + f"{gamma}\\mu Sv\\cdot m^2 / Ci\\cdot hr \\times {activity} Ci \\times \\frac{{1}}{{({cfg['dist']})^2}} \\times e^{{-0.693({cfg['exp_terms']})}} = {cfg['dose_str']}\\mu Sv / hr" + r'$'
        
        fig = plt.figure(figsize=(15, 1.5))
        plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
        fig.text(0.5, 0.5, eq, fontsize=22, ha='center', va='center', color='black')
        plt.axis('off')
        
        out_path = os.path.join(output_dir, cfg["filename"])
        plt.savefig(out_path, bbox_inches='tight', pad_inches=0.1, dpi=200)
        plt.close(fig)
        
        generated_paths.append(out_path)

    # Calculate max dose rate for final conclusion paragraph
    max_dose = max(dose_ab, dose_cd)
    max_dose_str = format_scientific(max_dose)
    
    eq_max = r'$' + max_dose_str + r'\times 10' # wait, format_scientific returns string without 10^ if not scientific!
    # Let's just reuse format_scientific which returns something like 4.42 \times 10^{-5}
    eq_max = r'$' + max_dose_str + r'$'
    
    fig = plt.figure(figsize=(1.5, 0.4))
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    fig.text(0.5, 0.5, eq_max, fontsize=16, ha='center', va='center', color='black')
    plt.axis('off')
    
    out_path_max = os.path.join(output_dir, "max_dose.jpg")
    plt.savefig(out_path_max, bbox_inches='tight', pad_inches=0.05, dpi=200)
    plt.close(fig)
    
    generated_paths.append(out_path_max)

    # Satisfaction text
    satisfaction = {
        "만족여부1": "만족" if dose_ab <= 10 else "불만족",
        "만족여부2": "만족" if dose_ab <= 1 else "불만족",
        "만족여부3": "만족" if dose_cd <= 10 else "불만족",
        "만족여부4": "만족" if dose_cd <= 1 else "불만족",
        "방사선조사상수": str(gamma),
        "가로거리": d_ab_str,
        "세로거리": d_cd_str,
    }

    return generated_paths, satisfaction

if __name__ == "__main__":
    # Test
    paths, sats = generate_math_images("Se-75", 60, 11, 0.8, 12, 1, 983, 45, 2000, 2000, 10000, 1000, "temp_math_imgs")
    for p in paths:
        print(p)
    print(sats)
