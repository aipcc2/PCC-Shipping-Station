\# Contributing to PCC Shipping Station



\## Rules of the road



1\. \*\*Never push directly to `main`.\*\* Every change goes through a branch and a pull request (PR).

2\. \*\*Every PR gets one review\*\* before merging.

3\. \*\*Test before opening a PR.\*\* Run the app from source, exercise the change you made.

4\. \*\*Commit messages should explain WHY, not just WHAT.\*\* "Fix typo" is bad. "Fix: COC ship date wasn't sanitized, caused crash on PO numbers with slashes" is good.



\## Workflow for a change (using GitHub Desktop)



1\. \*\*Pull latest from `main`:\*\*

&#x20;  In GitHub Desktop, make sure "Current Branch" is `main`, then click "Fetch origin" → "Pull origin".



2\. \*\*Create a new branch:\*\*

&#x20;  - Current Branch dropdown → "New Branch"

&#x20;  - Name it descriptively: `feature/csv-import`, `bugfix/draft-photo-missing`, `refactor/split-dialogs-module`



3\. \*\*Make your changes\*\* in your editor. Test by running `py -3.12 PCC\_Shipping.py`.



4\. \*\*Commit:\*\*

&#x20;  - GitHub Desktop shows changed files in the left panel.

&#x20;  - Write a clear commit message (summary + optional body).

&#x20;  - Click "Commit to <branch-name>".



5\. \*\*Push:\*\*

&#x20;  - Click "Publish branch" (first push) or "Push origin" (subsequent pushes).



6\. \*\*Open a Pull Request:\*\*

&#x20;  - GitHub Desktop will prompt you, OR go to the repo on github.com → "Pull requests" → "New pull request" → select your branch.

&#x20;  - Describe what changed and why. Mention any testing you did.

&#x20;  - Request a review from one teammate.



7\. \*\*Address review feedback\*\* if any, then merge once approved.



\## Branch naming conventions



\- `feature/<name>` — new functionality

\- `bugfix/<name>` — fixing broken behavior

\- `refactor/<name>` — cleanup without behavior change

\- `hotfix/<name>` — urgent production fix



\## Release process



1\. After merging changes that should go to production, on `main`:

2\. Update the version number near the top of `PCC\_Shipping.py` (e.g. `APP\_VERSION = "1.5.0"`).

3\. Commit with message `Release v1.5.0`.

4\. Create a Git tag: in GitHub Desktop → Repository → Create Tag → `v1.5.0`.

5\. Build the EXE: `build\_exe.bat`.

6\. Copy `dist\\PCC Shipping.exe` to the network share.

7\. Verify on at least one shipping PC before announcing.



\## When in doubt — ask



Drop questions into \[discussions / Slack channel / team chat — fill in your team's channel here]. Better to clarify before merging than fix in production.

