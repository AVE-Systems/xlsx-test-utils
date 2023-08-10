pipeline {
    agent any

    stages {
        stage('Build') {
            steps {
                sh '''
                    docker compose build php
                    docker compose run --rm php php -d memory_limit=-1 /usr/bin/composer install
                '''
            }
        }
        stage('PHPUnit') {
            steps {
                sh '''
                    docker compose run --rm php vendor/bin/phpunit tests
                '''
            }
        }
        stage('CSFixer') {
            steps {
                sh '''
                  docker-compose run --rm php vendor/bin/php-cs-fixer fix src --rules=@PhpCsFixer'
                  docker-compose run --rm php vendor/bin/php-cs-fixer fix tests --rules=@PhpCsFixer
                '''
            }
        }
    }
    post {
        always {
            sh 'docker compose down --volumes'
        }
    }
}
