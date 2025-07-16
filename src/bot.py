import asyncio
from dataclasses import dataclass, asdict
import json
import os
import sys
import traceback
from typing import Generic, TypeVar


from botbuilder.core import Middleware, TurnContext, MessageFactory, CardFactory, MemoryStorage
from botbuilder.schema import ActivityTypes, ChannelAccount
from teams import Application, ApplicationOptions, TeamsAdapter
from teams.ai import AIOptions
from teams.ai.models import AzureOpenAIModelOptions, OpenAIModel, OpenAIModelOptions
from teams.ai.planners import ActionPlanner, ActionPlannerOptions
from teams.ai.prompts import PromptManager, PromptManagerOptions
from teams.ai.actions import ActionTypes
from teams.state import TurnState
from teams.feedback_loop_data import FeedbackLoopData
from teams.ai.actions import ActionTypes, ActionTurnContext

from azure_ai_search_data_source import AzureAISearchDataSource, AzureAISearchDataSourceOptions
from custom_say_command import say_command
from config import Config

config = Config()

# Create AI components
model: OpenAIModel

model = OpenAIModel(
    AzureOpenAIModelOptions(
        api_key=config.AZURE_OPENAI_API_KEY,
        default_model=config.AZURE_OPENAI_MODEL_DEPLOYMENT_NAME,
        endpoint=config.AZURE_OPENAI_ENDPOINT,
    )
)
    
prompts = PromptManager(PromptManagerOptions(prompts_folder=f"{os.getcwd()}/prompts"))

# prompts.add_data_source(
#     AzureAISearchDataSource(
#         AzureAISearchDataSourceOptions(
#             name='azure-ai-search',
#             indexName='saeshav-test-teams-agent',
#             azureAISearchApiKey=config.AZURE_SEARCH_KEY,
#             azureAISearchEndpoint=config.AZURE_SEARCH_ENDPOINT,
#         )
#     )
# )
prompts.add_data_source(
    AzureAISearchDataSource(
        AzureAISearchDataSourceOptions(
            name='azure-ai-search-hr',
            indexName='saeshav-hr-index',
            azureAISearchApiKey=config.AZURE_SEARCH_KEY,
            azureAISearchEndpoint=config.AZURE_SEARCH_ENDPOINT,
        )
    )
)

prompts.add_data_source(
    AzureAISearchDataSource(
        AzureAISearchDataSourceOptions(
            name='azure-ai-search-procurement',
            indexName='saeshav-procurement-index',
            azureAISearchApiKey=config.AZURE_SEARCH_KEY,
            azureAISearchEndpoint=config.AZURE_SEARCH_ENDPOINT,
        )
    )
)


planner = ActionPlanner(
    ActionPlannerOptions(model=model, prompts=prompts, default_prompt="chat")
)

# Define storage and application
storage = MemoryStorage()
bot_app = Application[TurnState](
    ApplicationOptions(
        bot_app_id=config.APP_ID,
        storage=storage,
        adapter=TeamsAdapter(config),
        ai=AIOptions(planner=planner, enable_feedback_loop=True),
    )
)

@bot_app.ai.action(ActionTypes.SAY_COMMAND)
async def on_say(_context: ActionTurnContext, _state: TurnState):
    return await say_command(_context, _state, _context.data, feedback_loop_enabled=True)


card = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.4",
    "body": [
        { "type": "TextBlock", "text": "Pick an option:", "size": "Medium" }
    ],
    "actions": [
        { "type": "Action.Submit", "title": "HR", "data": { "choice": "A" } },
        { "type": "Action.Submit", "title": "Finances", "data": { "choice": "B" } }
    ]
}

# @bot_app.message("")
# async def on_message(_context: TurnContext, state: TurnState):
#     # If it's a card submit
#     if _context.activity.value:
#         data = _context.activity.value
#         choice = data.get("choice")
#         if choice == "A":
#             await _context.send_activity("✅ You clicked HR!")
#         elif choice == "B":
#             await _context.send_activity("✅ You clicked Finances!")
#         else:
#             await _context.send_activity(f"❓ Unknown choice: {choice}")
#         return True

#     return False

class TeamsMiddleware(Middleware):
    async def on_turn(self, context: TurnContext, next_call):
        # Detect new members joining
        if context.activity.type == ActivityTypes.conversation_update:
            if context.activity.members_added:
                for member in context.activity.members_added:
                    if member.id != context.activity.recipient.id:
                        await context.send_activity(
                            MessageFactory.attachment(CardFactory.adaptive_card(card))
                        )
            return  # stop after welcome

        # Detect Adaptive Card submit (button click)
        if context.activity.type == ActivityTypes.message and context.activity.value:
            data = context.activity.value
            choice = data.get("choice")
            if choice == "A":
                await context.send_activity("✅ You clicked HR!")
            elif choice == "B":
                await context.send_activity("✅ You clicked Finances!")
            else:
                await context.send_activity(f"❓ Unknown choice: {choice}")
            return  # stop after handling button

        # Otherwise let it continue to other handlers (like AI)
        await next_call()


# Attach this middleware to your adapter
bot_app.adapter.use(TeamsMiddleware())

@bot_app.error
async def on_error(context: TurnContext, error: Exception):
    # This check writes out errors to console log .vs. app insights.
    # NOTE: In production environment, you should consider logging this to Azure
    #       application insights.
    print(f"\n [on_turn_error] unhandled error: {error}", file=sys.stderr)
    traceback.print_exc()

    # Send a message to the user
    await context.send_activity("The agent encountered an error or bug.")

@bot_app.feedback_loop()
async def feedback_loop(_context: TurnContext, _state: TurnState, feedback_loop_data: FeedbackLoopData):
    # Add custom feedback process logic here.
    print(f"Your feedback is:\n{json.dumps(asdict(feedback_loop_data), indent=4)}")