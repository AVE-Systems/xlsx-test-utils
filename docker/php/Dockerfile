FROM php:7.2-cli

COPY --from=composer:1.10 /usr/bin/composer /usr/bin/composer
RUN composer --version

RUN apt-get update -y

RUN DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends apt-utils

RUN pecl install xdebug-2.6.1 && docker-php-ext-enable xdebug

RUN DEBIAN_FRONTEND=noninteractive apt-get install -y libpng-dev libjpeg62-turbo-dev

RUN docker-php-ext-install gd \
    && docker-php-ext-configure gd --with-jpeg-dir=/usr/include/

RUN DEBIAN_FRONTEND=noninteractive apt-get install -y libzip-dev zip \
    && docker-php-ext-configure zip --with-libzip && docker-php-ext-install zip
