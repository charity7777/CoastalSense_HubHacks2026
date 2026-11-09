# CoastalSense

**RBC HubHacks 2026 | Coastal Action Challenge**  
*Offline, AI-assisted analysis of underwater survey footage*

CoastalSense is a proof-of-concept application designed to support the preliminary review of underwater fish-survey footage. Built with Python, Streamlit, and the open-source YOLOv7 framework, the application combines AI-assisted fish detection, optional image enhancement, reviewable MeanCount estimates, and downloadable Excel reports.

The aim is to **support marine researchers, not replace their judgment**. Footage is processed locally rather than uploaded to a hosted AI service.

## Key Features

- **Underwater image and video analysis** — upload JPG, PNG, or MP4 files for local processing.
- **AI-assisted fish detection** — uses an externally trained N-MARINE YOLOv7 model to propose fish classifications and detections.
- **Visibility enhancement** — optional CLAHE-based contrast enhancement for challenging underwater imagery.
- **MeanCount estimation** — calculates a preliminary average fish count across systematically sampled video frames, including zero-detection frames.
- **Human review** — allows researchers to correct sample counts before exporting results.
- **Excel reporting** — generates a two-sheet report with an overview and sampled counts.


## Development & Experimentation

Before developing the final CoastalSense application, we conducted several experiments to evaluate how effectively YOLOv7 could detect and classify fish in underwater footage.

Our experiments included:

- **Initial fish detection:** Tested the pretrained N-MARINE YOLOv7 model on different fish images and underwater video footage.

- **CLAHE enhancement:** Applied Contrast Limited Adaptive Histogram Equalization (CLAHE) to improve visibility in underwater images and compared detection results before and after enhancement.

- **Image denoising:** Experimented with noise reduction combined with CLAHE to assess whether it could further improve image quality and detection.

- **Model evaluation:** Identified limitations, including missed fish detections, low-confidence predictions, and incorrect species classifications.

- **Approach selection:** Prioritized CLAHE for the prototype due to its simplicity and observed improvements in selected test images, while recognizing that enhancement does not guarantee better detection.

These experiments informed the development of CoastalSense, particularly the integration of visibility enhancement, preliminary MeanCount calculations, and optional manual corrections to support researcher verification.

**Note:** Experimental images and video footage obtained from third-party sources have not been included in this repository due to redistribution and copyright considerations.


## Model Parameters File — Important Notice

**The original parameters file has been intentionally excluded from this repository out of respect for its ownership and to avoid redistributing third-party materials without explicit authorization.**

The complete project source code is provided for inspection, evaluation, and further development. However, a suitable, authorized model-weights file is needed to perform AI detection.

In this project, the relevant file is **`best.pt`**: a PyTorch checkpoint containing pretrained N-MARINE YOLOv7 model weights. The model was trained and published by **Whidden Lab**.

**Official model source:** [Whidden Lab — N-MARINE Baseline Classifiers](https://huggingface.co/WhiddenLab/N-MARINE_baseline_classifiers)

### If You Wish to Run the Application

1. Clone or download this repository.
2. Obtain a **model checkpoint you are authorized to use**, such as one available from the official publisher at the link above.
3. Place the checkpoint in the repository's main folder, next to `coastalsense_app_rbc_simple_review.py`, and name it `best.pt`.
4. If you choose to run the application, follow the optional manual installation instructions below. **No model weights or automated setup/launch scripts are provided in this repository.**

The N-MARINE publisher provides more than one checkpoint. A different authorized checkpoint can produce results different from those shown in our demonstration.

The team's original model-weights file has **not been modified, shared, or included** in this public repository. This approach is intended to respect third-party ownership while allowing reviewers to understand the code and its intended workflow. 

## How CoastalSense Works

1. **Upload:** A researcher provides underwater footage or an image.
2. **Prepare:** The system samples frames and can apply contrast enhancement where appropriate.
3. **Detect:** YOLOv7 proposes object detections and species classifications using the locally available checkpoint.
4. **Summarize:** CoastalSense displays preliminary detections and a MeanCount-style average across sampled frames.
5. **Review:** A researcher may adjust sample counts to account for missed or incorrect detections.
6. **Export:** Results can be downloaded as an Excel report.

## Optional: Run Locally on Windows

**The repository is submitted primarily for source-code review. Running the application is not necessary to inspect our implementation.** For reviewers who want to run it, the steps below describe a manual local setup. No `setup.bat` or `start.bat` is included.

**Prerequisites:** Windows, Python **3.11 (64-bit)**, internet access for downloading Python packages, and an authorized N-MARINE-compatible `best.pt` checkpoint. The inference code is configured to run on the CPU; processing may be slow.

1. Download or clone the repository and open its main folder.
2. Obtain an authorized N-MARINE YOLOv7 checkpoint from the [official publisher](https://huggingface.co/WhiddenLab/N-MARINE_baseline_classifiers), rename it `best.pt` if necessary, and place it alongside `coastalsense_app_rbc_simple_review.py` and `detect.py`.
3. Open **PowerShell** in that folder and create a local Python environment:

   ```powershell
   py -3.11 -m venv .venv
   .\.venv\Scripts\python.exe -m pip install --upgrade pip
   ```

4. Install the CPU-compatible PyTorch packages and the application dependencies:

   ```powershell
   .\.venv\Scripts\python.exe -m pip install torch==2.3.1+cpu torchvision==0.18.1+cpu --index-url https://download.pytorch.org/whl/cpu
   .\.venv\Scripts\python.exe -m pip install -r requirements-coastalsense.txt
   ```

5. Launch CoastalSense:

   ```powershell
   .\.venv\Scripts\python.exe -m streamlit run coastalsense_app_rbc_simple_review.py --server.address 127.0.0.1 --server.port 8505
   ```

6. Open **http://127.0.0.1:8505** if needed, upload media you are authorized to use, and begin the analysis. Stop the app with **Ctrl+C** in the PowerShell terminal.

> **Dependency note:** `requirements-coastalsense.txt` records the compatibility versions used for the submission environment. The separate YOLOv7 `requirements.txt` is an upstream reference and should **not** be installed unchanged for this app: its original NumPy constraint conflicts with the compatible version used here (`numpy==1.26.4`). The environment was tested on the team's Windows laptop, but we cannot guarantee identical behavior on every computer.

## Project Structure

```text
CoastalSense-HubHacks2026/
├── coastalsense_app_rbc_simple_review.py           # Streamlit application
├── detect.py                                       # YOLOv7 inference script
├── models/                       		    # YOLOv7 model code
├── utils/                        		    # YOLOv7 support functions
├── cfg/                           		    # YOLOv7 model configurations
├── data/                          		    # YOLOv7 configuration assets
├── requirements-coastalsense.txt  		    # App-specific dependency versions
├── requirements.txt               		    # Original YOLOv7 dependency reference
├── README.md                      		    # Project documentation
├── THIRD_PARTY_NOTICES.md         		    # Upstream acknowledgements
└── LICENSE.md                     		    # GPL-3.0 licence text
```

`best.pt`, local virtual environments, third-party underwater footage and generated analysis outputs are intentionally excluded. This is a **code-review submission**; an authorized model checkpoint is needed for live inference.

## Prototype Scope and Limitations

CoastalSense is a **proof of concept**, not a validated species-monitoring system. The current prototype has the following limitations:

- Its detections and species predictions may be incorrect, particularly on unfamiliar underwater footage.
- **MeanCount is preliminary** and represents an average across sampled frames, not a count of unique individual fish.
- Automated tracking and individual-fish re-identification (for example, ByteTrack) have **not** been implemented.
- The draft-protocol option samples still images at selected timestamps. It does **not** automate every aspect of Coastal Action's draft video-analysis protocol, such as complete one-second observations, surrounding-context review, or independent reviewer quality assurance.
- Findings should be checked by qualified marine researchers before they are used for scientific conclusions.
- Deployment by Coastal Action would require its own policy approval, testing, and validation.

Future work could include model adaptation on appropriately authorized local footage, a watchlist of unexpected species, object tracking, and evaluation of acoustic imaging for low-visibility conditions.

## Data Privacy and Responsible Use

The application is designed for **local processing**. Video and images are not sent to a public AI inference service by CoastalSense. Manual installation downloads Python packages; the app does **not** upload user footage or model inputs to a hosted AI service.

Underwater footage from Memorial University, Coastal Action, or other third parties is **not included** unless redistribution has been explicitly authorized. For background on an original example source, see [Memorial University's Hawke Channel research page](https://sites.google.com/mun.ca/mcaproject/2023/leg-1-hawke-channel).

## Credits and Third-Party Materials

- **CoastalSense:** RBC HubHacks 2026 project developed in response to Coastal Action's monitoring challenge.
- **YOLOv7:** [WongKinYiu/yolov7](https://github.com/WongKinYiu/yolov7). The included YOLOv7 source code is subject to the original **GPL-3.0** licence. Preserve applicable notices and copyright information.
- **N-MARINE pretrained model:** [Whidden Lab — N-MARINE Baseline Classifiers](https://huggingface.co/WhiddenLab/N-MARINE_baseline_classifiers). The model was not trained or fine-tuned by this project team, and its weights are not bundled.
- **N-MARINE dataset:** Morris, C. J. et al. (2025), *Newfoundland Marine Refuge Fish Classification Dataset (N-Marine)*, [Government of Canada Open Data Portal](https://open.canada.ca/data/en/dataset/2ae46860-f82a-4127-bb1f-b02e36ef6a70).

See [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md) for additional attribution and reuse notes.

## Licence

The repository includes GPL-3.0-licensed YOLOv7 code and its corresponding licence document (`LICENSE.md`). **This does not grant any additional rights to redistribute third-party model weights or research footage.**
