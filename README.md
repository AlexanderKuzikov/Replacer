<p align="center">
  <a href="https://nodejs.org"><img alt="Node.js" src="https://img.shields.io/badge/Node.js-339933?logo=node.js&logoColor=white"></a>
  <a href="https://github.com/AlexanderKuzikov/Replacer/blob/main/LICENSE"><img alt="License" src="https://img.shields.io/badge/License-Apache--2.0-blue"></a>
</p>

<h1 align="center">Replacer</h1>
<p align="center">Очистка структур полей в DOCX-шаблонах</p>

---

Утилита для очистки структур полей (field structures) в DOCX-шаблонах. Работа с XML через jsdom, архивная обработка через adm-zip.

- **Очистка полей** — удаление служебных структур полей из XML.
- **DOCX** — прямая работа с архивной структурой документа.
- **jsdom** — парсинг и модификация XML-содержимого.

## Быстрый старт

```bash
git clone https://github.com/AlexanderKuzikov/Replacer.git
cd Replacer
npm install
node index.js
```

## Документация

- [`docs/CONTEXT.md`](docs/CONTEXT.md) — состояние проекта
- [`docs/DECISIONS.md`](docs/DECISIONS.md) — архитектурные решения

## Статус

**v1.0.0** — работает.

## Лицензия

[Apache-2.0](LICENSE) © Alexander Kuzikov
