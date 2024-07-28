"""
Criado em: 28-07
Por: Caique Rezende

Objetivo:
Este script tem como objetivo transferir compromissos da agenda do outlook, para a agenda do google
"""

## Carrega funções e libs
import os
from functions.extrai import extrai
from functions.insere import insere

## 1. Extrai os compromissos do outlook
extrai()

## 2. Insere os compromissos no gmail
insere()
