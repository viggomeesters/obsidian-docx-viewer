# Contributing

Keep changes aligned with the read-only inspection scope.

Before opening a pull request:

```bash
npm install
npm run build
npx tsc --noEmit
npm test
```

Do not add network services, cloud conversion, external app launch, macro execution, embedded object execution, clipboard access, or write-back support without a separate design review.
