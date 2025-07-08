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
                bat '"C:\\Users\\adm.luiz.vinicius\\AppData\\Local\\Programs\\Python\\Python312\\Scripts\\pip.exe" pip install -r requirements.txt'
            }
        }

        stage('Executar script Python') {
            steps {
                bat '"C:\\Users\\adm.luiz.vinicius\\AppData\\Local\\Programs\\Python\\Python312\\python.exe" NPS_ANUAL.py'
            }
        }
    }
}
