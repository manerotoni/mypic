# !!! THIS PROJECT MOVED TO https://github.com/manerotoni/mypic/ !!!

# Microscopy Pipeline Constructor (MyPiC)


MyPiC is a Visual Basic for Application (VBA) macro to be used with Zeiss confocal microscopes running with the ZEN software (version black). The macro has been developed in the group of Jan Ellenberg, EMBL, Heidelberg and replaces the AutofocusScreenMacro.

> **Supplementary software 2** of 
>>  Quantitative mapping of fluorescently tagged cellular proteins using FCS-calibrated four-dimensional imaging. Antonio Z Politi, Yin Cai, Nike Walther , M Julius Hossain ,
Birgit Koch, Malte Wachsmuth , Jan Ellenberg. (2018). Nature Protocols, Vol. 13,  DOI: 10.1038/nprot.2018.040
>
> Plase cite the above work when using this tool.

The macro allows  

* Autofocus based on reflection and fluorescence multi-location time series
* Fluorescence based tracking using center of mass of fluorescence signal
* Multi-location time-lapse experiments
* Flexible combination of several independent Z-stack and channel settings
* Flexible combination of several fluorescence correlation spectroscopy (FCS) measurements settings
* Adaptive Feedback microscopy support of two triggable imaging and FCS workflows

Please refer to the WiKi for further explanations 
and examples.

To concatenate files generated from a time lapse experiment in MyPiC refer the concat_utils. The original repository is in [https://git.embl.de/politi/concat](https://git.embl.de/politi/concat).
For adaptive feedback microscopy experiments you can use the ImageJ tool Automted FCS found in [https://git.embl.de/politi/adaptive_feedback_mic_fiji](https://git.embl.de/politi/adaptive_feedback_mic_fiji).


> **Disclaimer:**
> MyPiC for ZEN has been tested on Zeiss LSM 780 microscopes with ZEN 2010, 2011, and 2012, and LSM880 microscopes with ZEN2.1 and ZEN2.3. We don’t guarantee that it will work on other configurations and we don’t take any responsibility for damage occuring during or after use of MyPiC.

## Authors  
MyPiC has been originally developed at EMBL Heidelberg in the group of Jan Ellenberg

mail@apoliti.de, apoliti@mpibpc.mpg.de




