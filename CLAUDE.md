# Claude AI Assistant Instructions

## Project: Google Sheets to Code Converter

### Automatic Code Quality Checks

After making any changes to source code files in this project, you MUST:

1. **Run linting** and auto-fix any issues:
   ```bash
   npm run lint:fix
   ```
   If issues remain that can't be auto-fixed, manually fix them.

2. **Run formatting** and auto-fix any issues:
   ```bash
   npm run format:fix
   ```

3. **Run TypeScript type checking**:
   ```bash
   npm run typecheck
   ```
   Fix any type errors that appear.

4. **Run tests** to ensure nothing broke:
   ```bash
   npm test
   ```
   Fix any failing tests.

### Order of Operations

When editing source code:
1. Make the requested changes
2. Run `npm run lint:fix` 
3. Run `npm run format:fix`
4. Run `npm run typecheck` and fix any errors
5. Run `npm test` and fix any failures
6. Only report completion after all checks pass

### Quick Check Command

You can run all checks at once with:
```bash
npm run typecheck && npm run lint && npm run format && npm test
```

### Auto-fix Command Sequence

For automatic fixing, run in this order:
```bash
npm run format:fix && npm run lint:fix
```

### Important Notes

- **Always** run these checks after editing any `.ts`, `.js`, `.json` files in the `src/` directory
- **Always** ensure all checks pass before considering a task complete
- If a test fails after your changes, investigate why and fix it
- If you can't auto-fix an issue, manually resolve it
- The project uses Biome for linting and formatting
- The project has 121 tests that should all pass

### Project-Specific Context

- This project converts Google Sheets formulas to TypeScript/Python code
- It supports both OAuth2 and Service Account authentication
- It automatically detects and fetches referenced sheets
- Named ranges are automatically resolved
- Generated code includes CLI execution support

### File Structure

- Source code: `src/`
- Tests: `src/tests/`
- CLI: `src/cli/`
- Generators: `src/generators/`
- Parser: `src/parsers/`
- Utilities: `src/utils/`

### Common Commands

- `npm run cli -- convert --url <URL> --input-tabs "sheet1" --output-tabs "sheet2" --language typescript --output-file output.ts`
- `npm run build` - Compile TypeScript
- `npm run dev` - Run in watch mode
- `npm run test:watch` - Run tests in watch mode
- `npm run check` - Run Biome check (lint + format)
- `npm run check:fix` - Fix all Biome issues

### When Adding New Features

1. Write tests first (TDD approach when possible)
2. Implement the feature
3. Ensure all quality checks pass
4. Update TODO.md if completing a planned feature
5. Update README.md if adding user-facing features