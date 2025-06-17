#!/usr/bin/env bash
# exit on error
set -o errexit

# Instala as dependências
npm install

# Compila o TypeScript
npm run build

# Executa as migrations já com os arquivos transpilados
npm run typeorm:migrate
