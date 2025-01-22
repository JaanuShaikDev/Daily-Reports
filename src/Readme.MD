# Create environment 
conda create -p env python=3.12.4 -y

# Activate environment
conda activate ./env

# Install requirements
pip install -r requirements.txt

# Execute code
python ./src/DailyReports/DailyReports.py