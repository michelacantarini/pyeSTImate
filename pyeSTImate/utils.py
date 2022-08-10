import _collections_abc
import json
import math
import numpy as np


class Params:
    def __init__(self, json_path):
        with open(json_path) as f:
            params = json.load(f)
            self.__dict__.update(params)

    def save(self, json_path):
        with open(json_path, 'w') as f:
            params = json.dump(self.__dict__, f, indent=4)

    def update(self, json_path):
        with open(json_path) as f:
            params = json.load(f)
            self.__dict__.update(params)

    @property
    def dict(self):
        return self.__dict__


class CustomMaterials:
    def __init__(self, json_path):
        with open(json_path) as j:
            materials = json.load(j)
            self.__dict__.update(materials)

    def save(self, json_path):
        with open(json_path, 'w') as j:
            materials = json.dump(self.__dict__, j, indent=4)

    def update(self, json_path):
        with open(json_path) as j:
            materials = json.load(j)
            self.__dict__.update(materials)

    @property
    def dict(self):
        return self.__dict__


def avg_absorption_coefficient(alpha1, S1,
                               alpha2=0, S2=0,
                               alpha3=0, S3=0,
                               alpha4=0, S4=0,
                               alpha5=0, S5=0,
                               alpha6=0, S6=0,
                               alpha7=0, S7=0,
                               alpha8=0, S8=0,
                               alpha9=0, S9=0,
                               alpha10=0, S10=0):
    A1 = np.dot(alpha1, S1)
    A2 = np.dot(alpha2, S2)
    A3 = np.dot(alpha3, S3)
    A4 = np.dot(alpha4, S4)
    A5 = np.dot(alpha5, S5)
    A6 = np.dot(alpha6, S6)
    A7 = np.dot(alpha7, S7)
    A8 = np.dot(alpha8, S8)
    A9 = np.dot(alpha9, S9)
    A10 = np.dot(alpha10, S10)
    Atot = (A1 + A2 + A3 + A4 + A5 + A6 + A7 + A8 + A9 + A10)
    Stot = (S1 + S2 + S3 + S4 + S5 + S6 + S7 + S8 + S9 + S10)
    alpha_mean = Atot / Stot
    return alpha_mean


def modulation_depth_reduction_factor(r, Qf, rcf, F, Tf, Lsf, Lnf):
    A = (Qf / r**2) + (1 / rcf**2) * (1 + ((2 * math.pi * F * Tf) / 13.8)**2)**-1
    # print("A", A)
    B = ((2 * math.pi * F * Tf) / (13.8 * rcf**2)) * (1 + (2 * math.pi * F * Tf / 13.8)**2)**-1
    # print("B", B)
    C = (Qf / r**2) + (1 / rcf**2) + Qf * 10**((-Lsf + Lnf) / 10)
    # print("C", C)
    mfF = (math.sqrt(A**2 + B**2)) / C
    return mfF


def Speech_Transmssion_Index(MTI, alpha, beta):
    MTI_alpha_pond = np.multiply(MTI, alpha)
    MTI_beta_pond = np.multiply(MTI, beta)
    sum_MTI_alpha_pond = np.sum(MTI_alpha_pond)
    sum_MTI_beta_pond = np.sum(MTI_beta_pond)
    STI = sum_MTI_alpha_pond - sum_MTI_beta_pond
    return STI
