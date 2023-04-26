pipeline {
    agent any

    stages {
        stage('Verify Branch') {
            steps {
               echo "$GIT_BRANCH"
            }
        }
         stage('build')
        {
            steps
            {
                bat 'python --version'
                bat 'curl -sSL https://bootstrap.pypa.io/get-pip3.py -o get-pip3.py'
                bat 'python get-pip3.py'
                bat 'pip install -r requirements.txt'
                echo 'build phase has finished'
            }
        }

        stage('launch app fram') {
            steps {
                bat 'python main_fram.py'
            }
        }
    }
}