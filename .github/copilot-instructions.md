# Copilot Coding Agent Instructions

## Pull Request Title

- Set the PR title **once** at the start of a session to describe the **full scope** of the work (e.g. "Add overview.md with links to VPX script files").
- **Never** update the PR title on subsequent `report_progress` calls or when making additional commits.
- The PR title must reflect the overall goal of the branch, not the most recent commit message.

## Pull Request Description

- Use the `prDescription` field in `report_progress` to maintain a markdown checklist that tracks the overall plan and progress.
- Update the checklist items as work is completed (- [ ] → - [x]).
- Keep the checklist stable between updates; do not rewrite it from scratch on every call.

## Commit Messages

- Write clear, concise commit messages describing what changed in that specific commit.
- Commit messages are for git history, **not** for the PR title.
