.DEFAULT_GOAL := help
help:
	@grep -E '^[a-zA-Z-]+:.*?## .*$$' Makefile | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "[32m%-17s[0m %s\n", $$1, $$2}'
.PHONY: help

composer-install:
	docker-compose run --rm php composer install

run-tests: ## Run all PHPUnit tests
	docker-compose run --rm php vendor/bin/phpunit tests

fix-code: ## Format code style
	docker-compose run --rm php vendor/bin/php-cs-fixer fix src --rules=@PhpCsFixer
	docker-compose run --rm php vendor/bin/php-cs-fixer fix tests --rules=@PhpCsFixer
