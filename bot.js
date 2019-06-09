// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, CardFactory, MessageFactory  } = require('botbuilder');
const axios = require('axios');
const addressParser = require('parse-address');
const validator = require('validator');


class EchoBot extends ActivityHandler {
    constructor(conversationState, userState) {
        super();

        // defining step state for the conversation
        this.step = {
            name: 'ask name',
            address: 'ask address',
            email: 'ask email',
            book: 'ask book',
            pickingBook: 'pick a book',
            confirmBook: 'confirm a book',
            summary: 'summary',
            quit: 'quit'
        }

        // The accessor names for the conversation data and user profile state property
        const CONVERSATION_DATA_PROPERTY = 'conversationData';
        const USER_PROFILE_PROPERTY = 'userProfile';
        // Create the state property accessors for the conversation data and use
        this.conversationData = conversationState.createProperty(CONVERSATION_DATA_PROPERTY);
        this.userProfile = userState.createProperty(USER_PROFILE_PROPERTY);

        // The state management objects for the conversation and user state.
        this.conversationState = conversationState;
        this.userState = userState;

        // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
        this.onMessage(async (turnContext, next) => {

            // Get the state properties from the turn context.
            const userProfile = await this.userProfile.get(turnContext, {});

            // get the conversation state with initalized data
            const conversationData = await this.conversationData.get(turnContext, { step: this.step.name });

            await this.fillOutOrder(userProfile, conversationData, turnContext);
            
            // await context.sendActivity(`You said '${ context.activity.text }'`);
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onDialog(async (turnContext, next) => {
            // Save any state changes. The load happened during the execution of the Dialog.
            await this.conversationState.saveChanges(turnContext, false);
            await this.userState.saveChanges(turnContext, false);
        
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity('Hello and welcome!');
                }
            }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

    async fillOutOrder(user, conversationState, turnContext) {
        switch(conversationState.step) {
            case this.step.name:
                await turnContext.sendActivity('What is your name?');

                // set new step for the bot
                conversationState.step = this.step.address;
                break;
            case this.step.address:
                if (!this.isValidName(turnContext.activity.text)) {
                    await turnContext.sendActivity(`Your name can not contain number or empty. It is impossible!!`);
                    await turnContext.sendActivity('Please provide your name again!!');
                    return;
                }
                // save user name from last question
                user.name = turnContext.activity.text;

                // affirm user name
                await turnContext.sendActivity(`Oh, hey ${ user.name }. Nice to see you here!`);
                // ask for address
                await turnContext.sendActivity('What is your shipping address? Please provide number and street. Please provide carefully include street address, city, state and zipcode.');

                // set new step for the bot
                conversationState.step = this.step.email;
                break;
            case this.step.email:
                const address = turnContext.activity.text;
                // check if user address is valid
                if (!this.isValidAddress(address)) {
                    await turnContext.sendActivity(`Please provide your valid address. It must contain street address, (apt), city, two letter state and zipcode.`);
                    return;
                }

                user.address = address;
                
                await turnContext.sendActivity(`Please provide your email.`);

                conversationState.step = this.step.book;
                break;
            case this.step.book:
                const email = turnContext.activity.text;
                // check user email is valid
                if (!validator.isEmail(email)) {
                    await turnContext.sendActivity('Please provide valid email.');
                    return;
                }

                user.email = email;

                await turnContext.sendActivity(`Please provide your book name.`);

                conversationState.step = this.step.pickingBook;
                break;
            case this.step.pickingBook:
                    const book = turnContext.activity.text.trim();
                    const { data: { items } } = await axios.get(`https://www.googleapis.com/books/v1/volumes?q=${book}`);
                    const { title, volumeInfo: { imageLinks : { thumbnail } } } = items[0];
                    
                    const card = CardFactory.heroCard(
                        title,
                        'A beautiful book',
                        [thumbnail],
                        ['buy']
                    );
                    const message = MessageFactory.attachment(card);

                    await turnContext.sendActivity(message);
                    await turnContext.sendActivity('Is this the book you are looking for? Click Buy if it is correct. Other wise type "No"');

                    conversationState.step = this.step.confirmBook;
                    conversationState.pickingBook = book.charAt(0).toUpperCase() + book.slice(1);
                break;
            case this.step.confirmBook:
                    const answer = turnContext.activity.text.toLowerCase();
                    if (answer === 'buy') {
                        user.book = conversationState.pickingBook;

                        await turnContext.sendActivity('Okay, awesome!! Are you ready?');

                        conversationState.step = this.step.summary;
                    } else if (answer === 'no') {
                        await turnContext.sendActivity('Please provide your book name again.');
                        conversationState.step = this.step.pickingBook;
                    } else {
                        await turnContext.sendActivity('Sorry, I dont get it? Please answer again.');
                    }
                    break;
            case this.step.summary:
                
                await turnContext.sendActivity(`Thank you for choosing my service. Summary of your info:\n${user.name}\n${user.address}\n${user.email}\n${user.book}`);
                await turnContext.sendActivity('Thank you for shopping. Confirmation email will be sent to you shortly');

                await axios.post('https://hc-mailing-to-customer.herokuapp.com/email', {
                    name: user.name,
                    email: user.email,
                    content: `Your order confirmation:\n${user.name}\n${user.address}\n${user.book}`
                })

                
                // reset step
                user.name = '';
                user.email = '';
                user.address = '';
                user.book = '';
                conversationState.step = this.step.name;
                break;
        }
    }

    isValidName(userName) {
        const name = userName.trim();

        // if customer name contain digit or empty string 
        if (/\d+/g.test(name) || !name) {
            return false;
        }

        return true;
    }

    isValidAddress(address) {
        const parsedAddress = addressParser.parseLocation(address);

        // if user provide an address
        if (!parsedAddress) {
            return false
        }
        // if user provide valid address
        if (!parsedAddress.number || !parsedAddress.street || !parsedAddress.zip) {
            return false;
        }
        
        return true;
    }
}

module.exports.EchoBot = EchoBot;
