// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, CardFactory } = require('botbuilder');
const { QnAMaker } = require('botbuilder-ai');
const WelcomeCard = require('./services/WelcomeCard.json');
const CADreamAct = require('./services/CADreamAct.json');
const FinAidCard = require('./services/FinancialAid.json');
const contactCard = require('./services/contactCard.json');
const Colleges = require('./services/Colleges.json');
const DisabilityCard = require('./services/DisabilityCard.json');

class QnABot extends ActivityHandler {
    /**
     * @param {any} logger object for logging events, defaults to console if none is provided
     */
    constructor(logger) {
        super();
        if (!logger) {
            logger = console;
            logger.log('[QnaMakerBot]: logger not passed in, defaulting to console');
        }

        try {
            var endpointHostName = process.env.QnAEndpointHostName;
            if (!endpointHostName.startsWith('https://')) {
                // eslint-disable-next-line no-mixed-spaces-and-tabs,no-tabs
		    endpointHostName = 'https://' + endpointHostName;
            }

            if (!endpointHostName.endsWith('/qnamaker')) {
                // eslint-disable-next-line no-mixed-spaces-and-tabs,no-tabs
		    endpointHostName = endpointHostName + '/qnamaker';
            } this.qnaMaker = new QnAMaker({
                knowledgeBaseId: process.env.QnAKnowledgebaseId,
                endpointKey: process.env.QnAAuthKey,
                host: endpointHostName
            });
        } catch (err) {
            logger.warn(`QnAMaker Exception: ${ err } Check your QnAMaker configuration in .env`);
        }
        this.logger = logger;

        // If a new user is added to the conversation, send them a greeting message
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity({ attachments: [CardFactory.adaptiveCard(WelcomeCard)] });
                }
            }

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        // When a user sends a message, perform a call to the QnA Maker service to retrieve matching Question and Answer pairs.
        this.onMessage(async (context, next) => {
            this.logger.log('Calling QnA Maker');

            const qnaResults = await this.qnaMaker.getAnswers(context);

            if (qnaResults[0].answer === 'Above you can see information regarding the California Dream Act.') {
                await context.sendActivity({ attachments: [CardFactory.adaptiveCard(CADreamAct)] });
            }

            if (qnaResults[0].answer === 'Here are the steps to applying for Financial Aid:') {
                await context.sendActivity({ attachments: [CardFactory.adaptiveCard(FinAidCard)] });
            }

            if (qnaResults[0].answer === 'Above you can see information for contacting various school departments.') {
                await context.sendActivity({ attachments: [CardFactory.adaptiveCard(contactCard)] });
            }
            if (qnaResults[0].answer === 'Above are links to the different colleges that make up csudh') {
                await context.sendActivity({ attachments: [CardFactory.adaptiveCard(Colleges)] });
            }
            if (qnaResults[0].answer === 'Above you will find information on the resources available at CSUDH for students with disabilities.') {
                await context.sendActivity({ attachments: [CardFactory.adaptiveCard(DisabilityCard)] });
            }
            // If an answer was received from QnA Maker, send the answer back to the user.
            if (qnaResults[0]) {
                await context.sendActivity(qnaResults[0].answer);
                // eslint-disable-next-line brace-style
            }
            // If no answers were returned from QnA Maker, reply with help.
            else {
                await context.sendActivity('Sorry, no answers were found for your question, please try again.');
            }

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }
}

module.exports.QnABot = QnABot;
