# Monthly Presentation Generator

A Streamlit-based application that transforms PDF or text input into professional PowerPoint presentations using AWS Bedrock for content analysis and Plotly for dynamic chart generation.

## Features
- **Input Options**: Upload PDF files or enter text directly.
- **Content Types**: Supports logistics reports, business reports, and narrative stories.
- **Dynamic Charts**: Generates bar, line, pie, and table visualizations based on input data.
- **Multilingual Support**: Translates content to 13 languages using AWS Bedrock.
- **Customizable Styles**: Choose from Professional, Creative, Minimalist, or Corporate themes.
- **AWS Integration**: Uses Bedrock for intelligent content analysis and translation.

## Project Structure
monhly_customized/ 
├── app.py                  # Main application code 
├── .env                    # AWS credentials 
├── uploads/                # Temporary storage for uploaded files 
├── output/                 # Generated presentations 
├── user_templates/         # Custom template storage 
├── templates/              # Default templates and styles 
│   ├── styles/             # Theme JSON files 
│   ├── icons/              # Decorative images/icons 
├── app.log                 # Application logs 
├── requirements.txt        # Dependencies


## Installation & Setup 

## Prerequisites 

- Python 3.8 or higher. 

- AWS Bedrock access with valid credentials. 

- Internet connection (for URL-based templates). 

- Git (optional, for cloning the repository). 

## Steps to Install 

- Clone the Repository (optional): 

git clone https://github.com/your-repo/smart-presentation-generator.git 
cd monhly_customized
Alternatively, save the provided code as app.py. 

- Create a Virtual Environment (recommended): 

python -m venv venv 
source venv/bin/activate  # On Windows: venv\Scripts\activate 
  

- Install Dependencies: Create a requirements.txt with the listed dependencies and run: 

pip install -r requirements.txt 
  

- Set Up AWS Credentials: Create a .env file in the project root with: 

AWS_ACCESS_KEY_ID=your_access_key_id 
AWS_SECRET_ACCESS_KEY=your_secret_access_key 
AWS_SESSION_TOKEN=your_session_token  # Optional 
AWS_DEFAULT_REGION=us-east-1 
BEDROCK_MODEL_ID=anthropic.claude-3-sonnet-20240229-v1:0 
BEDROCK_REGION=us-east-1 
  

- Verify AWS setup: 

aws sts get-caller-identity 
  

- Run the Application: 

streamlit run app.py 
  

- Access the app at http://localhost:8501. 