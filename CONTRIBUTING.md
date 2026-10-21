# Contributing to czWebview

Thank you for your interest in contributing to czWebview! This document provides guidelines for contributing to this project.

## How to Contribute

### Reporting Bugs
1. Check the [existing issues](https://github.com/cyberzilla/czWebview/issues) to avoid duplicates
2. Create a new issue with:
   - Clear, descriptive title
   - Steps to reproduce the bug
   - Expected behavior vs actual behavior
   - VB6 version, Windows version, and WebView2 Runtime version
   - Code snippet that demonstrates the issue

### Suggesting Features
1. Open an issue with the `enhancement` label
2. Describe the feature and its use case
3. If possible, reference the relevant WebView2 API documentation

### Submitting Code Changes

#### Prerequisites
- VB6 IDE (SP6 recommended)
- WebView2 Runtime installed
- Windows SDK with MIDL compiler (only if modifying the type library)

#### Workflow
1. Fork the repository
2. Create a feature branch: `git checkout -b feature/my-feature`
3. Make your changes
4. Test with the demo projects (SimpleTest, WhatsappBot, CalculatorWebApp)
5. Commit with descriptive messages: `git commit -m "feat: add XYZ support"`
6. Push to your fork: `git push origin feature/my-feature`
7. Open a Pull Request

#### Commit Message Convention
We follow [Conventional Commits](https://www.conventionalcommits.org/):
- `feat:` — New feature
- `fix:` — Bug fix
- `docs:` — Documentation changes
- `refactor:` — Code refactoring (no functional change)
- `demo:` — Changes to demo projects
- `chore:` — Build scripts, maintenance tasks

### Code Style Guidelines
- Always use `Option Explicit`
- Prefix private methods with `pv`
- Prefix member variables with `m_`
- Prefix Friend event sinks with `Sink`
- Use `On Error GoTo EH` pattern for error handling
- Add section headers with `'=========` dividers
- Use 4-space indentation

## Project Structure
```
src/czWebview.ctl           — Public API (properties, methods, events)
src/czWebviewCallback.cls   — Internal COM callback handler
typelib/czWebview.odl       — Interface definitions
typelib/czWebview.tlb       — Compiled type library
external/WebView2Loader.dll — Microsoft x86 loader
demo/                       — Demo projects
```

## Questions?
Open an issue with the `question` label or start a discussion.
