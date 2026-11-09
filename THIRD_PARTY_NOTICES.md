# Attribution and third-party materials

## YOLOv7 source code

The YOLOv7 inference code (`detect.py`, `models/`, `utils/` and related
configuration) comes from [WongKinYiu/yolov7](https://github.com/WongKinYiu/yolov7),
which is licensed under GPL-3.0. Keep the upstream copyright notices and
licence when including that code. `LICENSE.md` in this repository provides the
GPL-3.0 text; do not remove applicable original notices.

## N-MARINE model

The `best.pt` used by this application is a model trained and released by
Whidden Lab, **not** a checkpoint fine-tuned by the CoastalSense team.
Official model source:
https://huggingface.co/WhiddenLab/N-MARINE_baseline_classifiers

No explicit permission to redistribute the model weights was established
from the publisher's model card. For that reason, `best.pt` is not bundled
in this repository. Reviewers who wish to run the application can obtain an
authorized N-MARINE checkpoint directly from its official publisher.
This is not a statement that the weights are GPL-3.0 licensed.

The official repository contains multiple checkpoints corresponding to
different train/test splits. A checkpoint other than the one used for the
hackathon demo may produce different results.

## Video and images

Do not upload Memorial University underwater footage, Coastal Action recordings,
or other third-party media without verifying redistribution permission.
For the Memorial University Hawke Channel video, link to the original source:
https://sites.google.com/mun.ca/mcaproject/2023/leg-1-hawke-channel

## Attribution

Morris, C. J., Ayyagari, K. D., Porter, D., Nguyen, Q. K., Hanlon, J., &
Whidden, C. (2025). *Newfoundland Marine Refuge Fish Classification Dataset
(N-Marine)*. Government of Canada Open Data Portal.
https://open.canada.ca/data/en/dataset/2ae46860-f82a-4127-bb1f-b02e36ef6a70
