# PDF Parser

# Installation

1. Clone the Repository

```
git clone https://github.com/your-username/your-repository-name.git
cd your-repository-name
```

2. Set Up a Virtual Environment 

```
python -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`
```

3. Install Required Packages

Install all the required packages using pip.

```
pip install -r requirements.txt
```

4. Environment Variables
The project uses environment variables to manage configuration settings. Follow these steps to set up the environment variables:

- Create a .env file: In the root directory of the project, create a file named .env.
- Add the required environment variables: Open the .env file in a text editor and add the following lines:

```
CLIENT_ID=your_client_id_here
CLIENT_SECRET=your_client_secret_here
TENANT_ID=your_tenant_id_here
DRIVE_ID=your_drive_id_here
EXCEL_FILE_PATH=your_excel_file_path_here
SECRET_KEY=your_secret_key_here
```

- Replace your_client_id_here, your_client_secret_here, and so on with the actual values for your application.
- Update your Django settings: Ensure your Django project is configured to read from the .env file. You may use libraries like python-decouple or django-environ for this purpose.


5. Running the Server

```
python manage.py migrate  # Run database migrations
python manage.py runserver  # Start the development server
```

6. Accessing the Application
Open a web browser and navigate to http://localhost:8000/