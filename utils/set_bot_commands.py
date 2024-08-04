from aiogram import types


async def set_default_commands(dp):
    await dp.bot.set_my_commands(
        [
            types.BotCommand("start", "Start the bot. After the command you can send excel file."),
            types.BotCommand("help", "Guidance for using bot!\n"
                                     " After /start command you need to send file,\n"
                                     " because after the command bot will be in state,\n"
                                     " it expect only file from user. \n"
                                     "If you send text or other thing bot won't reply you."),
        ]
    )
