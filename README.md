# pyeSTImate
A python-based tool for the *speech transmission index* prediction.
## Description
PyeSTImate is a predictive tool of speech transmission index in room acoustic environments, representing an extension of the room impulse response (RIR) simulator included in the software package *pyroomacoustics* [1]. It aims to calculate the speech transmission index (STI) of a room without recourse to direct measurements of the RIRs, which are simulated based on the room geometry, materials, furniture, type of occupants, and positions of source and receivers. After calculating the reverberation time (RT) at the listening points, the average STI is determined by the indirect method according to UNI 11532-1:2018 [2] and IEC 60268-16:2020 [3], and the evaluation procedures of UNI 11532-2:2020 [4].\
PyeSTImate consists of two main blocks: the first, based on *pyroomacoustics*, returns simulated room impulse responses and reverberation times in octave bands that feed into the second part of the algorithm, designed to calculate the STI using the indirect method. Based on the average STI value, the tool also provides the intelligibility rating, a qualification according to a five-point scale of the quality of speech intelligibility. The block diagram of the tool is shown in Fig. 1.

![Alt text](/imgs/block_diagram.png?raw=true)
<p align="center">
<i>
Fig. 1 - Block diagram of pyeSTImate
</i>
</p>

## Usage
Instructions for use can be found in [pyeSTimate User Guide](/docs/user_guide.pdf).

## Dependencies
PyeSTImate has been implemented in python. The main dependencies are:
* matplotlib (https://matplotlib.org/)
* numpy (https://numpy.org/)
* pandas (https://pandas.pydata.org/)
* pyroomacoustics (https://pyroomacoustics.readthedocs.io/en/pypi-release/)
* scikit-learn (https://scikit-learn.org/stable/index.html)

## Example
Will be added soon.

## Contribute
All contributions to improving the speed and usability of the code are welcome!

## References
[1] Scheibler, Robin, Eric Bezzam, and Ivan Dokmanić. "Pyroomacoustics: A python package for audio room simulation and array processing algorithms." 2018 IEEE international conference on acoustics, speech and signal processing (ICASSP). IEEE, 2018.\
[2] UNI 11532-1:2018. "Caratteristiche acustiche interne di ambienti confinati - Metodi di progettazione e tecniche di valutazione - Parte 1: Requisiti generali."\
[3] IEC 60268-16:2020. "Sound system equipment - Part 16: Objective rating of speech intelligibility by speech transmission index."\
[4] UNI 11532-2:2020. "Caratteristiche acustiche interne di ambienti confinati - Metodi di progettazione e tecniche di valutazione - Parte 2: Settore scolastico."
