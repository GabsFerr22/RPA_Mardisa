pipeline {
    agent any

    stages {
        stage('Clonar o repositório') {
            steps {
                git branch: 'main',
                    url: 'https://github.com/GabsFerr22/RPA_Mardisa.git'
            }
        }

        stage('Instalar dependências') {
            steps {
                bat 'pip install -r requirements.txt'
            }
        }

        stage('Executar script Python') {
            steps {
                bat 'python NPS_ANUAL.py'
            }
        }
    }
}
