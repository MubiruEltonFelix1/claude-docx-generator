# Contributing to DOCX Generator

Thanks for taking the time to improve this tool. It's small and focused, and contributions should stay that way — useful, targeted, and easy to review.

---

## How to contribute

**Found a bug?** [Open an issue](https://github.com/MubiruEltonFelix1/claude-docx-generator/issues) with a description of what went wrong, the script you pasted, and the error you got.

**Want a feature?** Open an issue first and describe the use case. Once the approach is agreed on, implementation is straightforward.

**Ready to code?** Follow the steps below.

---

## Development setup

```bash
git clone https://github.com/MubiruEltonFelix1/claude-docx-generator.git
cd claude-docx-generator
npm install
```

Test that everything works:

```bash
node generate2.js
# → outputs/MIT_Masterclass_Geometry_of_Meaning.docx should appear
```

---

## Making a change

1. Fork the repo and create a branch:
   ```bash
   git checkout -b feat/stream-interception
   ```

2. Make your change. Keep the scope tight — one feature or one fix per PR.

3. Test manually by pasting a `docx` script into `paste_here.js` and verifying the output lands in `outputs/`.

4. Open a pull request with a clear title and a one-paragraph description of what changed and why.

---

## Areas where help is wanted

| Feature | Notes |
|---|---|
| Stream interception | Catch `fs.createWriteStream` paths for scripts that stream instead of using Buffer writes |
| CLI `--output` flag | Let users specify a custom output directory at run time |
| DOCX preview | Open the generated file automatically after writing |
| Better error messages | Friendly output when the pasted script has a syntax error or missing dependency |

If you want to work on one of these, comment on the relevant issue or open a new one.

---

## Code style

- No linter is configured — just match the style of the existing files.
- Keep `runner.js` focused: it patches `fs` and nothing else.
- Comments are welcome when the behaviour isn't obvious.

---

## License

By contributing, you agree that your changes will be licensed under the [MIT License](./LICENSE).
