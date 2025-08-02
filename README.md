
## Update Log
- [2025/08/1] code and README.


### Requirements
- Python 3.6+
- OpenCV (`cv2`)
- NumPy
- SciPy
- scikit-image
- Matplotlib
- openpyxl

### How to use
1:Prepare watermark and cover images
	Place watermark images such as HDU_W_20.png in the root directory.
	Include your test cover images (e.g., peppers.bmp, baboon.bmp) in the same folder.

2:Run the script
```
python code.py
```  

3:Check outputs
Results will be saved in folders like __Try/
Includes visualizations, extracted watermark images and logs
PSNR, SSIM, NC, BER recorded in Try_.xlsx

4:Attack Modes
	The following distortion attacks are applied during robustness evaluation:
	JPEG compression (Q=70, 50, 30)
	Image scaling (0.5×, 0.7×, 1.5×)
	Gaussian noise (σ=0.01, 0.03)
	Salt & pepper noise (amount=0.001, 0.01)
	Rotation
	Mean and median blurring
	Combined attacks (e.g., scale + noise + JPEG)

## Citation
Please cite the paper if you use the code.

