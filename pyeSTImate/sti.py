import argparse
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pyroomacoustics as pra
from utils import *


"""
Speech Transmission Index expressing the quality of speech transmission without amplification system relative to
an acoustic path between a speaker and a listener computed with the indirect method according to UNI 11532-1:2018
and IEC 60268-16:2020.
"""


# Link to config file
parser = argparse.ArgumentParser()
parser.add_argument('--config_path', type=str)
parser.add_argument('--materials_path', type=str)
args = parser.parse_args()
params = Params(args.config_path)
materials = CustomMaterials(args.materials_path)

# Link to annotations file
annotations_file = params.annotations_file
df = pd.read_excel(annotations_file)
print(df.head())


# Iterate over all the dataframe rows
for index, row in df.iterrows():
    # ----------------------------------------------------------------------------------------------- #
    '''Parameters'''
    # ----------------------------------------------------------------------------------------------- #
    room = row['Room']
    length = row['length']  # m
    depth = row['depth']  # m
    height = row['height']  # m
    # ----------------------------------------------------------------------------------------------- #
    s01_floor = row['area_floor_material_1']  # m^2
    s02_floor = row['area_floor_material_2']  # m^2
    s03_floor = row['area_floor_material_3']  # m^2
    s04_floor = row['area_floor_material_4']  # m^2
    s05_floor = row['area_floor_material_5']  # m^2

    s01_pers = row['area_Apers_1']  # m^2
    s02_pers = row['area_Apers_2']  # m^2
    s03_pers = row['area_Apers_3']  # m^2
    s04_pers = row['area_Apers_4']  # m^2
    s05_pers = row['area_Apers_5']  # m^2
    # ----------------------------------------------------------------------------------------------- #
    s01_ceiling = row['area_ceiling_material_1']  # m^2
    s02_ceiling = row['area_ceiling_material_2']  # m^2
    s03_ceiling = row['area_ceiling_material_3']  # m^2
    s04_ceiling = row['area_ceiling_material_4']  # m^2
    s05_ceiling = row['area_ceiling_material_5']  # m^2
    s_object_ceiling = row['area_Aobj_11']  # m^2
    # ----------------------------------------------------------------------------------------------- #
    s01_wall_S = row['area_south_wall_material_1']  # m^2
    s02_wall_S = row['area_south_wall_material_2']  # m^2
    s03_wall_S = row['area_south_wall_material_3']  # m^2
    s04_wall_S = row['area_south_wall_material_4']  # m^2
    s05_wall_S = row['area_south_wall_material_5']  # m^2

    s01_object_wall_S = row['area_south_wall_Aobj_1']  # m^2
    s02_object_wall_S = row['area_south_wall_Aobj_2']  # m^2
    s03_object_wall_S = row['area_south_wall_Aobj_3']  # m^2
    s04_object_wall_S = row['area_south_wall_Aobj_4']  # m^2
    s05_object_wall_S = row['area_south_wall_Aobj_5']  # m^2
    # ----------------------------------------------------------------------------------------------- #
    s01_wall_E = row['area_east_wall_material_1']  # m^2
    s02_wall_E = row['area_east_wall_material_2']  # m^2
    s03_wall_E = row['area_east_wall_material_3']  # m^2
    s04_wall_E = row['area_east_wall_material_4']  # m^2
    s05_wall_E = row['area_east_wall_material_5']  # m^2

    s01_object_wall_E = row['area_east_wall_Aobj_1']  # m^2
    s02_object_wall_E = row['area_east_wall_Aobj_2']  # m^2
    s03_object_wall_E = row['area_east_wall_Aobj_3']  # m^2
    s04_object_wall_E = row['area_east_wall_Aobj_4']  # m^2
    s05_object_wall_E = row['area_east_wall_Aobj_5']  # m^2
    # ----------------------------------------------------------------------------------------------- #
    s01_wall_N = row['area_north_wall_material_1']  # m^2
    s02_wall_N = row['area_north_wall_material_2']  # m^2
    s03_wall_N = row['area_north_wall_material_3']  # m^2
    s04_wall_N = row['area_north_wall_material_4']  # m^2
    s05_wall_N = row['area_north_wall_material_5']  # m^2

    s01_object_wall_N = row['area_north_wall_Aobj_1']  # m^2
    s02_object_wall_N = row['area_north_wall_Aobj_2']  # m^2
    s03_object_wall_N = row['area_north_wall_Aobj_3']  # m^2
    s04_object_wall_N = row['area_north_wall_Aobj_4']  # m^2
    s05_object_wall_N = row['area_north_wall_Aobj_5']  # m^2
    # ----------------------------------------------------------------------------------------------- #
    s01_wall_W = row['area_west_wall_material_1']  # m^2
    s02_wall_W = row['area_west_wall_material_2']  # m^2
    s03_wall_W = row['area_west_wall_material_3']  # m^2
    s04_wall_W = row['area_west_wall_material_4']  # m^2
    s05_wall_W = row['area_west_wall_material_5']  # m^2

    s01_object_wall_W = row['area_west_wall_Aobj_1']  # m^2
    s02_object_wall_W = row['area_west_wall_Aobj_2']  # m^2
    s03_object_wall_W = row['area_west_wall_Aobj_3']  # m^2
    s04_object_wall_W = row['area_west_wall_Aobj_4']  # m^2
    s05_object_wall_W = row['area_west_wall_Aobj_5']  # m^2
    # ----------------------------------------------------------------------------------------------- #
    center_freqs = [125.0, 250.0, 500.0, 1000.0, 2000.0, 4000.0, 8000.0]  # 7 octave bands (Hz)
    # ----------------------------------------------------------------------------------------------- #
    mat01_floor = row['floor_material_1']
    mat02_floor = row['floor_material_2']
    mat03_floor = row['floor_material_3']
    mat04_floor = row['floor_material_4']
    mat05_floor = row['floor_material_5']

    mat01_floor = getattr(materials, mat01_floor)
    mat02_floor = getattr(materials, mat02_floor)
    mat03_floor = getattr(materials, mat03_floor)
    mat04_floor = getattr(materials, mat04_floor)
    mat05_floor = getattr(materials, mat05_floor)
    # ----------------------------------------------------------------------------------------------- #
    mat01_ceiling = row['ceiling_material_1']
    mat02_ceiling = row['ceiling_material_2']
    mat03_ceiling = row['ceiling_material_3']
    mat04_ceiling = row['ceiling_material_4']
    mat05_ceiling = row['ceiling_material_5']

    mat01_ceiling = getattr(materials, mat01_ceiling)
    mat02_ceiling = getattr(materials, mat02_ceiling)
    mat03_ceiling = getattr(materials, mat03_ceiling)
    mat04_ceiling = getattr(materials, mat04_ceiling)
    mat05_ceiling = getattr(materials, mat05_ceiling)
    # ----------------------------------------------------------------------------------------------- #
    mat01_wall = row['wall_material_1']
    mat02_wall = row['wall_material_2']
    mat03_wall = row['wall_material_3']
    mat04_wall = row['wall_material_4']
    mat05_wall = row['wall_material_5']

    mat01_wall = getattr(materials, mat01_wall)
    mat02_wall = getattr(materials, mat02_wall)
    mat03_wall = getattr(materials, mat03_wall)
    mat04_wall = getattr(materials, mat04_wall)
    mat05_wall = getattr(materials, mat05_wall)
    # ----------------------------------------------------------------------------------------------- #
    pers01 = row['Apers_1']
    pers02 = row['Apers_2']
    pers03 = row['Apers_3']
    pers04 = row['Apers_4']
    pers05 = row['Apers_5']

    pers01 = getattr(materials, pers01)
    pers02 = getattr(materials, pers02)
    pers03 = getattr(materials, pers03)
    pers04 = getattr(materials, pers04)
    pers05 = getattr(materials, pers05)
    # ----------------------------------------------------------------------------------------------- #
    object01 = row['Aobj_1']
    object02 = row['Aobj_2']
    object03 = row['Aobj_3']
    object04 = row['Aobj_4']
    object05 = row['Aobj_5']
    object_ceiling = row['Aobj_11']

    object01 = getattr(materials, object01)
    object02 = getattr(materials, object02)
    object03 = getattr(materials, object03)
    object04 = getattr(materials, object04)
    object05 = getattr(materials, object05)
    object_ceiling = getattr(materials, object_ceiling)
    # ----------------------------------------------------------------------------------------------- #
    floor_scattering = row['scattering_floor']
    ceiling_scattering = row['scattering_ceiling']
    wall_S_scattering= row['scattering_south_wall']
    wall_E_scattering = row['scattering_east_wall']
    wall_N_scattering = row['scattering_north_wall']
    wall_W_scattering = row['scattering_west_wall']
    # ----------------------------------------------------------------------------------------------- #
    method = params.simulation_method  # ISM or hybrid simulator
    fs = params.rir_sampling_rate  # sampling rate
    max_order = params.max_order_reflections  # reflections order
    decay_db = params.decay_db  # decay in decibels for which the time is estimated
    # ----------------------------------------------------------------------------------------------- #
    n_receivers = params.n_receivers  # number of receivers (min=1, max=4)
    source_coordinates = list(map(float, row['source_coordinates'].split(',')))
    mic01_coordinates = list(map(float, row['receiver_1_coordinates'].split(',')))
    mic02_coordinates = list(map(float, row['receiver_2_coordinates'].split(',')))
    mic03_coordinates = list(map(float, row['receiver_3_coordinates'].split(',')))
    mic04_coordinates = list(map(float, row['receiver_4_coordinates'].split(',')))
    # ----------------------------------------------------------------------------------------------- #
    Qf = 1  # source directivity factor
    c = 343.0  # sound speed (m/s)
    Lsf = [62.9, 62.9, 59.2, 53.2, 47.2, 41.2, 35.2]  # UNI EN ISO 9921:2014 Table H.2 (speech spectrum at 1 m in front of the mouth of a male speaker (dB))
    Lnf = [41.0, 43.0, 50.0, 47.0, 42.0, 42.0, 39.0]  # UNI EN ISO 9921:2014 Table H.1 (ambient noise spectrum (dB))
    F = [0.63, 0.8, 1, 1.25, 1.6, 2, 2.5, 3.15, 4, 5, 6.3, 8, 10, 12.5]  # BS EN 60268-16:2011 par.A.2.2 (STI modulation frequencies (Hz))
    alpha_male = [0.085, 0.127, 0.230, 0.233, 0.309, 0.224, 0.173]  # BS EN 60268-16:2011 Table A.3 (MTI octave band weighting factors)
    beta_male = [0.085, 0.078, 0.065, 0.011, 0.047, 0.095, 0]  # BS EN 60268-16:2011 Table A.3 (MTI octave band weighting factors)
    alpha_female = [0, 0.117, 0.223, 0.216, 0.328, 0.250, 0.194]  # BS EN 60268-16:2011 Table A.3 (MTI octave band weighting factors)
    beta_female = [0, 0.099, 0.066, 0.062, 0.025, 0.076, 0]  # BS EN 60268-16:2011 Table A.3 (MTI octave band weighting factors)
    # ----------------------------------------------------------------------------------------------- #
    speaker_gender = params.speaker_gender
    plot = params.plot
    print_recap = params.print_recap
    # ----------------------------------------------------------------------------------------------- #
    # Transformation from list to array
    mic_coordinates = np.array([mic01_coordinates, mic02_coordinates, mic03_coordinates, mic04_coordinates])
    Lsf = np.array(Lsf)
    Lnf = np.array(Lnf)
    # ----------------------------------------------------------------------------------------------- #
    '''Room dimensions and source level correction'''
    # ----------------------------------------------------------------------------------------------- #
    room_dim = [length, depth, height]
    V = length * depth * height  # Volume of the room (m^3)

    # Leq of the source depending on the room volume (UNI 11532-2:2020 prospect 11)
    if V < 250:
        Lsf = Lsf
    elif V >= 250:
        Lsf = Lsf + 10
    # ----------------------------------------------------------------------------------------------- #
    '''Distance source-receiver'''
    # ----------------------------------------------------------------------------------------------- #
    r = []
    for i in range(n_receivers):
        distance = np.linalg.norm(source_coordinates - mic_coordinates[i])
        r.append(distance)
    # ----------------------------------------------------------------------------------------------- #
    '''Recap input data'''
    # ----------------------------------------------------------------------------------------------- #
    if print_recap:
        print("""----------------------------------------
        \nINPUT DATA RECAP\n""")
        print("Room", room)
        print("Room dimensions (length, depth, and height in m):", length, depth, height)
        print("Volume of the room (m^3):", np.round(V, 2))
        print("Source coordinates (m):", source_coordinates)
        print("Number of receivers:", n_receivers)
        for i in range(n_receivers):
            print("Receiver {} coordinates (m):".format(i+1), mic_coordinates[i])
        print("Distances source-receivers (m):", np.round(r, 2))
        print("Ambient noise spectrum UNI EN ISO 9921:2014 Table H.1 (dB):", Lnf)
        print("Speech spectrum UNI EN ISO 9921:2014 Table H.2 (dB):", Lsf, "speaker_gender:", params.speaker_gender)

        print("Floor materials:")
        print(row['floor_material_1'], "=", np.round(length*depth-(s02_floor+s03_floor+s04_floor+s05_floor), 2), "m^2")
        if s02_floor != 0:
            print(row['floor_material_2'], "=", np.round(s02_floor, 2), "m^2")
        if s03_floor != 0:
            print(row['floor_material_3'], "=", np.round(s03_floor, 2), "m^2")
        if s04_floor != 0:
            print(row['floor_material_4'], "=", np.round(s04_floor, 2), "m^2")
        if s05_floor != 0:
            print(row['floor_material_5'], "=", np.round(s05_floor, 2), "m^2")

        print("Ceiling materials:")
        print(row['ceiling_material_1'], "=", np.round(length*depth-(s02_ceiling+s03_ceiling+s04_ceiling+s05_ceiling), 2), "m^2")
        if s02_ceiling != 0:
            print(row['ceiling_material_2'], "=", np.round(s02_ceiling, 2), "m^2")
        if s03_ceiling != 0:
            print(row['ceiling_material_3'], "=", np.round(s03_ceiling, 2), "m^2")
        if s04_ceiling != 0:
            print(row['ceiling_material_4'], "=", np.round(s04_ceiling, 2), "m^2")
        if s05_ceiling != 0:
            print(row['ceiling_material_5'], "=", np.round(s05_ceiling, 2), "m^2")

        print("Wall materials:")
        print(row['wall_material_1'], "=", np.round((2*(length+depth)*height)
                                                    - ((s02_wall_S + s02_wall_E + s02_wall_N + s02_wall_W)
                                                    + (s03_wall_S + s03_wall_E + s03_wall_N + s03_wall_W)
                                                    + (s04_wall_S + s04_wall_E + s04_wall_N + s04_wall_W)
                                                    + (s05_wall_S + s05_wall_E + s05_wall_N + s05_wall_W)), 2),
              "m^2")
        if s02_wall_S != 0 or s02_wall_E != 0 or s02_wall_N != 0 or s02_wall_W != 0:
            print(row['wall_material_2'], "=",  np.round((s02_wall_S + s02_wall_E + s02_wall_N + s02_wall_W), 2), "m^2")
        if s03_wall_S != 0 or s03_wall_E != 0 or s03_wall_N != 0 or s03_wall_W != 0:
            print(row['wall_material_3'], "=",  np.round((s03_wall_S + s03_wall_E + s03_wall_N + s03_wall_W), 2), "m^2")
        if s04_wall_S != 0 or s04_wall_E != 0 or s04_wall_N != 0 or s04_wall_W != 0:
            print(row['wall_material_4'], "=",  np.round((s04_wall_S + s04_wall_E + s04_wall_N + s04_wall_W), 2), "m^2")
        if s05_wall_S != 0 or s05_wall_E != 0 or s05_wall_N != 0 or s05_wall_W != 0:
            print(row['wall_material_5'], "=",  np.round((s05_wall_S + s05_wall_E + s05_wall_N + s05_wall_W), 2), "m^2")

        print("Occupants/Audience:")
        if s01_pers != 0:
            print(row['Apers_1'], "=", np.round(s01_pers, 2), "m^2")
        if s02_pers != 0:
            print(row['Apers_2'], "=", np.round(s02_pers, 2), "m^2")
        if s03_pers != 0:
            print(row['Apers_3'], "=", np.round(s03_pers, 2), "m^2")
        if s04_pers != 0:
            print(row['Apers_4'], "=", np.round(s04_pers, 2), "m^2")
        if s05_pers != 0:
            print(row['Apers_5'], "=", np.round(s05_pers, 2), "m^2")
        elif s01_pers == 0 and s02_pers == 0 and s03_pers == 0 and s04_pers == 0 and s05_pers == 0:
            print("None")

        print("Furniture:")
        if s01_object_wall_S != 0 or s01_object_wall_E != 0 or s01_object_wall_N != 0 or s01_object_wall_W != 0:
            print(row['Aobj_1'], "=", np.round((s01_object_wall_S + s01_object_wall_E + s01_object_wall_N + s01_object_wall_W), 2), "m^2")
        if s02_object_wall_S != 0 or s02_object_wall_E != 0 or s02_object_wall_N != 0 or s02_object_wall_W != 0:
            print(row['Aobj_2'], "=", np.round((s02_object_wall_S + s02_object_wall_E + s02_object_wall_N + s02_object_wall_W), 2), "m^2")
        if s03_object_wall_S != 0 or s03_object_wall_E != 0 or s03_object_wall_N != 0 or s03_object_wall_W != 0:
            print(row['Aobj_3'], "=", np.round((s03_object_wall_S + s03_object_wall_E + s03_object_wall_N + s03_object_wall_W), 2), "m^2")
        if s04_object_wall_S != 0 or s04_object_wall_E != 0 or s04_object_wall_N != 0 or s04_object_wall_W != 0:
            print(row['Aobj_4'], "=", np.round((s04_object_wall_S + s04_object_wall_E + s04_object_wall_N + s04_object_wall_W), 2), "m^2")
        if s05_object_wall_S != 0 or s05_object_wall_E != 0 or s05_object_wall_N != 0 or s05_object_wall_W != 0:
            print(row['Aobj_5'], "=", np.round((s05_object_wall_S + s05_object_wall_E + s05_object_wall_N + s05_object_wall_W), 2), "m^2")
        if s_object_ceiling != 0:
            print(row['Aobj_11'], "=", np.round(s_object_ceiling, 2), "m^2")
        elif s01_object_wall_S == 0 and s01_object_wall_E == 0 and s01_object_wall_N == 0 and s01_object_wall_W == 0 and s_object_ceiling == 0:
            print("None")
        print("Simulation method:", params.simulation_method)
        print("Max order reflections:", params.max_order_reflections)
        print("Reverberation Time Decay (dB):", params.decay_db)
        print("RIR sampling rate (Hz):", params.rir_sampling_rate)
        print("----------------------------------------")
        print("\nSIMULATION RESULTS")
    # ----------------------------------------------------------------------------------------------- #
    '''Average absorption coefficients'''
    # ----------------------------------------------------------------------------------------------- #
    # Floor
    if s01_floor < 0:
        raise Exception("Check exception surfaces: main area < 0")
    avg_alpha_floor = avg_absorption_coefficient(mat01_floor, s01_floor,
                                                 mat02_floor, s02_floor,
                                                 mat03_floor, s03_floor,
                                                 mat04_floor, s04_floor,
                                                 mat05_floor, s05_floor,
                                                 pers01, s01_pers,
                                                 pers02, s02_pers,
                                                 pers03, s03_pers,
                                                 pers04, s04_pers,
                                                 pers05, s05_pers)
    print("\nFloor average absorption coefficients:", np.round(avg_alpha_floor, 2))
    print("Floor scattering coefficient:", floor_scattering)
    # ----------------------------------------------------------------------------------------------- #
    # Ceiling
    if s01_ceiling < 0:
        raise Exception("Check exception surfaces: main area < 0")
    avg_alpha_ceiling = avg_absorption_coefficient(mat01_ceiling, s01_ceiling,
                                                   mat02_ceiling, s02_ceiling,
                                                   mat03_ceiling, s03_ceiling,
                                                   mat04_ceiling, s04_ceiling,
                                                   mat05_ceiling, s05_ceiling,
                                                   object_ceiling, s_object_ceiling)
    print("Ceiling average absorption coefficients:", np.round(avg_alpha_ceiling, 2))
    print("Ceiling scattering coefficient:", ceiling_scattering)
    # ----------------------------------------------------------------------------------------------- #
    # Wall_S - south
    if s01_wall_S < 0:
        raise Exception("Check exception surfaces: main area < 0")
    avg_alpha_wall_S = avg_absorption_coefficient(mat01_wall, s01_wall_S,
                                                  mat02_wall, s02_wall_S,
                                                  mat03_wall, s03_wall_S,
                                                  mat04_wall, s04_wall_S,
                                                  mat05_wall, s05_wall_S,
                                                  object01, s01_object_wall_S,
                                                  object02, s02_object_wall_S,
                                                  object03, s03_object_wall_S,
                                                  object04, s04_object_wall_S,
                                                  object05, s05_object_wall_S)
    print("South wall average absorption coefficients:", np.round(avg_alpha_wall_S, 2))
    print("South wall scattering coefficient:", wall_S_scattering)
    # ----------------------------------------------------------------------------------------------- #
    # Wall_E - east
    if s01_wall_E < 0:
        raise Exception("Check exception surfaces: main area < 0")
    avg_alpha_wall_E = avg_absorption_coefficient(mat01_wall, s01_wall_E,
                                                  mat02_wall, s02_wall_E,
                                                  mat03_wall, s03_wall_E,
                                                  mat04_wall, s04_wall_E,
                                                  mat05_wall, s05_wall_E,
                                                  object01, s01_object_wall_E,
                                                  object02, s02_object_wall_E,
                                                  object03, s03_object_wall_E,
                                                  object04, s04_object_wall_E,
                                                  object05, s05_object_wall_E)
    print("East wall average absorption coefficients:", np.round(avg_alpha_wall_E, 2))
    print("East wall scattering coefficient:", wall_E_scattering)
    # ----------------------------------------------------------------------------------------------- #
    # Wall_N - north
    if s01_wall_N < 0:
        raise Exception("Check exception surfaces: main area < 0")
    avg_alpha_wall_N = avg_absorption_coefficient(mat01_wall, s01_wall_N,
                                                  mat02_wall, s02_wall_N,
                                                  mat03_wall, s03_wall_N,
                                                  mat04_wall, s04_wall_N,
                                                  mat05_wall, s05_wall_N,
                                                  object01, s01_object_wall_N,
                                                  object02, s02_object_wall_N,
                                                  object03, s03_object_wall_N,
                                                  object04, s04_object_wall_N,
                                                  object05, s05_object_wall_N)
    print("North wall average absorption coefficients:", np.round(avg_alpha_wall_N, 2))
    print("North wall scattering coefficient:", wall_N_scattering)
    # ----------------------------------------------------------------------------------------------- #
    # Wall_W - west
    if s01_wall_W < 0:
        raise Exception("Check exception surfaces: main area < 0")
    avg_alpha_wall_W = avg_absorption_coefficient(mat01_wall, s01_wall_W,
                                                  mat02_wall, s02_wall_W,
                                                  mat03_wall, s03_wall_W,
                                                  mat04_wall, s04_wall_W,
                                                  mat05_wall, s05_wall_W,
                                                  object01, s01_object_wall_W,
                                                  object02, s02_object_wall_W,
                                                  object03, s03_object_wall_W,
                                                  object04, s04_object_wall_W,
                                                  object05, s05_object_wall_W)
    print("West wall average absorption coefficients:", np.round(avg_alpha_wall_W, 2))
    print("West wall scattering coefficient:", wall_W_scattering)
    # ----------------------------------------------------------------------------------------------- #
    '''Materials definition'''
    # ----------------------------------------------------------------------------------------------- #
    ceiling_absorption = {'description': 'ceiling_absorption', 'coeffs': avg_alpha_ceiling, 'center_freqs': center_freqs}
    floor_absorption = {'description': 'floor_absorption', 'coeffs': avg_alpha_floor, 'center_freqs': center_freqs}
    wall_S_absorption = {'description': 'wall_S_absorption', 'coeffs': avg_alpha_wall_S, 'center_freqs': center_freqs}
    wall_E_absorption = {'description': 'wall_E_absorption', 'coeffs': avg_alpha_wall_E, 'center_freqs': center_freqs}
    wall_N_absorption = {'description': 'wall_N_absorption', 'coeffs': avg_alpha_wall_N, 'center_freqs': center_freqs}
    wall_W_absorption = {'description': 'wall_W_absorption', 'coeffs': avg_alpha_wall_W, 'center_freqs': center_freqs}

    m = pra.make_materials(
        ceiling=(ceiling_absorption, ceiling_scattering),
        floor=(floor_absorption, floor_scattering),
        south=(wall_S_absorption, wall_S_scattering),
        east=(wall_E_absorption, wall_E_scattering),
        north=(wall_N_absorption, wall_N_scattering),
        west=(wall_W_absorption, wall_W_scattering)
    )
    # ----------------------------------------------------------------------------------------------- #
    '''Room simulation'''
    # ----------------------------------------------------------------------------------------------- #
    if method == 'ISM':
        room = pra.ShoeBox(room_dim,
                           fs=fs,
                           materials=m,
                           max_order=max_order,
                           air_absorption=False,
                           use_rand_ism=True,
                           max_rand_disp=0.5
                           )

    elif method == 'Hybrid':
        room = pra.ShoeBox(
            room_dim,
            fs=fs,
            materials=m,
            max_order=max_order,
            ray_tracing=True,
            air_absorption=True,
        )
        # Activate the ray tracing
        room.set_ray_tracing(receiver_radius=0.1, n_rays=100000, energy_thres=1e-7, time_thres=5)
    # ----------------------------------------------------------------------------------------------- #
    '''Source positioning'''
    # ----------------------------------------------------------------------------------------------- #
    # Add the sound source (talker)
    room.add_source(source_coordinates)
    # ----------------------------------------------------------------------------------------------- #
    '''STI computation'''
    # ----------------------------------------------------------------------------------------------- #
    # Definition of 7 octave bands starting from 125 Hz
    octave = pra.acoustics.OctaveBandsFactory(base_frequency=125, fs=fs, n_fft=512)
    n_octave_bands = len(octave.centers)

    t60_simul_list = []
    STI_simul_list = []
    for i in range(n_receivers):
        # Add two receivers (listeners)
        mic_array = np.c_[mic_coordinates[i]]
        room.add_microphone_array(mic_array)

    # Run the simulation (this will also build the RIR automatically)
    room.compute_rir()
    for i in range(n_receivers):
        rir = room.rir[i][0]
        print("\nRIR at position {}:\n".format(i+1), rir)

        # Reverberation Time (Schroeder)
        t60 = pra.experimental.measure_rt60(room.rir[i][0], decay_db=decay_db, fs=fs, plot=plot)
        print("\nRT60 at position {}:".format(i+1), np.round(t60 * 1000, 2), "ms")

        t60_simul_list.append(t60)
        t60_avg_simul = (sum(t60_simul_list)) / (len(t60_simul_list))

        # Reverberation Time in octave bands (Schroeder)
        rt = []
        for j in range(n_octave_bands):
            rir_octave_bands = octave.analysis(rir, band=j)
            rt_octave_bands = pra.experimental.rt60.measure_rt60(rir_octave_bands, fs=fs, decay_db=decay_db, plot=False, rt60_tgt=None)
            rt.append(rt_octave_bands)
        print("\nSchroeder RT60 in octave bands at position {} (s):\n".format(i+1), np.round(rt, 2))

        # Critical Distances (7)
        cost = (2 * V / c)
        rt = np.array(rt)
        rcf = np.sqrt(np.multiply(cost, rt))
        print("\nCritical distances at position {} (m):\n".format(i+1), np.round(rcf, 2))

        # Modulation Transfer Function matrix (14 x 7 = 98)
        MTF = []
        for k in range(len(F)):
            for l in range(len(rt)):
                mfF = modulation_depth_reduction_factor(r[i], Qf, rcf[l], F[k], rt[l], Lsf[l], Lnf[l])
                MTF.append(mfF)
        MTF = np.reshape(MTF, (14, 7))
        print("\nMTF calculated by Schroeder equation IEC 60268-16 at position {}:\n".format(i+1), np.round(MTF, 3))

        # Modulation Transfer Indexes (7)
        MTI = np.mean(MTF, axis=0)
        print("\nModulation Transfer Indexes at position {}:\n".format(i+1), np.round(MTI, 3))

        # Speech Transmission Index
        if speaker_gender == 'male':
            STI_simul = Speech_Transmssion_Index(MTI, alpha_male, beta_male)
        elif speaker_gender == 'female':
            STI_simul = Speech_Transmssion_Index(MTI, alpha_female, beta_female)
        STI_simul_list.append(STI_simul)

    STI_P1_simul = STI_simul_list[0]
    print("\nSTI P1 =", np.round(STI_P1_simul, 2))
    STI_P2_simul = STI_simul_list[1]
    print("STI P2 =", np.round(STI_P2_simul, 2))
    STI_P3_simul = STI_simul_list[2]
    print("STI P3 =", np.round(STI_P3_simul, 2))
    STI_P4_simul = STI_simul_list[3]
    print("STI P4 =", np.round(STI_P4_simul, 2))

    STI_avg_simul = (sum(STI_simul_list))/(len(STI_simul_list))

    # Speech comprehension quality
    if 0 < STI_avg_simul <= 0.3:
        IR_simul = 'bad'
        IR_label_simul = 1
        print("\nSTI average =", np.round(STI_avg_simul, 2), " --> Bad speech quality")
    elif 0.30 < STI_avg_simul <= 0.45:
        IR_simul = 'poor'
        IR_label_simul = 2
        print("\nSTI average =", np.round(STI_avg_simul, 2), " --> Poor speech quality")
    elif 0.45 < STI_avg_simul <= 0.60:
        IR_simul = 'fair'
        IR_label_simul = 3
        print("\nSTI average =", np.round(STI_avg_simul, 2), " --> Fair speech quality")
    elif 0.60 < STI_avg_simul <= 0.75:
        IR_simul = 'good'
        IR_label_simul = 4
        print("\nSTI average =", np.round(STI_avg_simul, 2), " --> Good speech quality")
    elif 0.75 < STI_avg_simul <= 1.00:
        IR_simul = 'excellent'
        IR_label_simul = 5
        print("\nSTI average =", np.round(STI_avg_simul, 2), " --> Excellent speech quality")

    # Plots
    if plot:
        # room.plot(img_order=1, mic_marker_size=0.01, no_axis=False)
        room.plot_rir(FD=False)
        fig = plt.gcf()
        fig.set_size_inches(20, 10)
        plt.show()


df_final = pd.concat([df[['Room',
                          'length',
                          'depth',
                          'height']],
                      pd.DataFrame({'V': [V],
                                    'T60_avg_simul': t60_avg_simul,
                                    'STI_P1_simul': STI_P1_simul,
                                    'STI_P2_simul': STI_P2_simul,
                                    'STI_P3_simul': STI_P3_simul,
                                    'STI_P4_simul': STI_P4_simul,
                                    'STI_avg_simul': STI_avg_simul,
                                    'IR_simul': IR_simul,
                                    'IR_label': IR_label_simul})],
                     axis=1)

df_final.to_excel("output.xlsx", sheet_name='STI')
print("\n", df_final.round(2))
