"""
Agent configuration: uses OpenAI Agents SDK with Qwen Plus (千问 Plus).
OPENAI_API_KEY environment variable must be set in system environment.
"""

import os
from pathlib import Path
from openai import AsyncOpenAI
from agents import Agent
from agents.models.openai_chatcompletions import OpenAIChatCompletionsModel

from selenium_tools import (
    fill_login_form,
    open_browser,
    login,
    login_and_create_public_group,
    login_and_create_new_person,
    login_and_create_new_department,
    select_org_structure_tool,
    create_public_group_tool,
    create_new_person_tool,
    create_new_department_tool,
    edit_person_tool,
    login_and_edit_person,
    get_current_page_info,
    navigate_to_login_page,
    set_language_to_simplified_chinese,
    take_screenshot,
    close_browser,
)

_SKILL_FILE = Path(__file__).with_name("skill.md")


def _load_system_prompt() -> str:
    """Load agent instructions from the editable Markdown skill file."""
    try:
        content = _SKILL_FILE.read_text(encoding="utf-8").strip()
    except FileNotFoundError as exc:
        raise RuntimeError(f"Missing skill file: {_SKILL_FILE}") from exc

    if not content:
        raise RuntimeError(f"Skill file is empty: {_SKILL_FILE}")
    return content


def create_agent() -> Agent:
    """Create and return the eTeams Passport test agent configured with Qwen Plus."""
    qwen_client = AsyncOpenAI(
        api_key=os.environ.get("OPENAI_API_KEY"),
        base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
    )

    model = OpenAIChatCompletionsModel(
        model="qwen-plus",
        openai_client=qwen_client,
    )

    agent = Agent(
        name="eTeams Passport 测试助手",
        model=model,
        instructions=_load_system_prompt(),
        tools=[
            open_browser,
            set_language_to_simplified_chinese,
            fill_login_form,
            login,
            login_and_create_public_group,
            login_and_create_new_person,
            login_and_create_new_department,
            select_org_structure_tool,
            create_public_group_tool,
            create_new_person_tool,
            create_new_department_tool,
            edit_person_tool,
            login_and_edit_person,
            get_current_page_info,
            navigate_to_login_page,
            take_screenshot,
            close_browser,
        ],
    )
    return agent
