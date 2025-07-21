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
from state_store import storage

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

with open(os.path.join(os.getcwd(), "indexers/folders.json")) as f:
    folder_index_map = json.load(f)

index_names = list(folder_index_map.values())

for index_name in index_names:
    prompts.add_data_source(
        AzureAISearchDataSource(
            AzureAISearchDataSourceOptions(
                name=index_name,
                indexName=index_name,
                azureAISearchApiKey=config.AZURE_SEARCH_KEY,
                azureAISearchEndpoint=config.AZURE_SEARCH_ENDPOINT,
            )
        )
    )



planner = ActionPlanner(
    ActionPlannerOptions(model=model, prompts=prompts, default_prompt="chat")
)

# Define storage and applicationurnStat
# storage = MemoryStorage()
bot_app = Application[TurnState](
    ApplicationOptions(
        bot_app_id=config.APP_ID,
        storage=storage,
        adapter=TeamsAdapter(config),
        ai=AIOptions(planner=planner, enable_feedback_loop=True),
    )
)

actions = []
for index_name in index_names:
    display_title = index_name.capitalize()
    actions.append({
        "type": "Action.Submit",
        "title": display_title,
        "data": { "choice": index_name }
    })

card = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.4",
    "body": [
        { "type": "TextBlock", "text": "Pick an option:", "size": "Medium" }
    ],
    "actions": actions
}


@bot_app.ai.action(ActionTypes.SAY_COMMAND)
async def on_say(_context: ActionTurnContext, _state: TurnState):
    return await say_command(_context, _state, _context.data, feedback_loop_enabled=True)


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
        # if context.activity.type == ActivityTypes.message and context.activity.value:
        #     data = context.activity.value
        #     choice = data.get("choice")

        #     user_id = context.activity.from_property.id
        #     conv_id = context.activity.conversation.id
        #     key = f"{conv_id}:{user_id}"

        #     if choice == "A":
        #         await storage.write({ key: {"selected_sources": "hr"} })
        #         await context.send_activity("✅ You clicked HR!")
        #     elif choice == "B":
        #         await storage.write({ key: {"selected_sources": "procurement"} })
        #         await context.send_activity("✅ You clicked Finances!")
                
        #     else:
        #         await context.send_activity(f"❓ Unknown choice: {choice}")
        #     return  # stop after handling button
        if context.activity.type == ActivityTypes.message and context.activity.value:
            data = context.activity.value
            choice = data.get("choice")

            user_id = context.activity.from_property.id
            conv_id = context.activity.conversation.id
            storage_key = f"{conv_id}:{user_id}"

            if choice in index_names:
                await storage.write({storage_key: {"selected_sources": choice}})
                await context.send_activity(f"✅ You selected {choice}!")
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