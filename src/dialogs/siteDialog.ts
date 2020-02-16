import { Attachment, AttachmentLayoutTypes, CardFactory } from 'botbuilder';

import {
    ComponentDialog,
    ConfirmPrompt,
    DialogTurnResult,
    TextPrompt,
    WaterfallDialog,
    WaterfallStepContext
} from 'botbuilder-dialogs';

import { AliasResolverDialog } from './aliasResolverDialog';
import { OwnerResolverDialog } from './ownerResolverDialog';
import { SiteDetails } from './siteDetails';

const TEXT_PROMPT = 'textPrompt';
const OWNER_RESOLVER_DIALOG = 'ownerResolverDialog';
const ALIAS_RESOLVER_DIALOG = 'aliasResolverDialog';
const CONFIRM_PROMPT = 'confirmPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

import GenericCard from '../resources/generic.json';
import SiteTypesCard from '../resources/siteTypes.json';
import SummaryCard from '../resources/summary.json';

export class SiteDialog extends ComponentDialog {
    constructor(id: string) {
        super(id || 'siteDialog');
        this
            .addDialog(new TextPrompt(TEXT_PROMPT))
            .addDialog(new OwnerResolverDialog(OWNER_RESOLVER_DIALOG))
            .addDialog(new AliasResolverDialog(ALIAS_RESOLVER_DIALOG))
            .addDialog(new ConfirmPrompt(CONFIRM_PROMPT))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.siteTypeStep.bind(this),
                this.titleStep.bind(this),
                this.descriptionStep.bind(this),
                this.ownerStep.bind(this),
                this.aliasStep.bind(this),
                this.confirmStep.bind(this),
                this.finalStep.bind(this)
            ]));
        this.initialDialogId = WATERFALL_DIALOG;
    }    

    /**
     * If a site type has not been provided, prompt for one.
     */
    private async siteTypeStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        if (!siteDetails.siteType) {

            const siteTypeCards: Attachment[] = SiteTypesCard.cards.map((card: any) => CardFactory.adaptiveCard(card));
            await stepContext.context.sendActivity({ 
                attachmentLayout: AttachmentLayoutTypes.Carousel, 
                attachments: siteTypeCards 
            });

            return await stepContext.prompt(TEXT_PROMPT, '');

        } else {
            return await stepContext.next(siteDetails.siteType);
        }
    }
    
    /**
     * If a title has not been provided, prompt for one.
     */
    private async titleStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        siteDetails.siteType = stepContext.result.value;

        if (!siteDetails.title) {

            const promptText = `Provide a title for your ${siteDetails.siteType} site`;
            const titleCard: Attachment = CardFactory.adaptiveCard(JSON.parse(
                JSON.stringify(GenericCard).replace('$Placeholder', promptText)));
            
            await stepContext.context.sendActivity({ attachments: [titleCard] });
            return await stepContext.prompt(TEXT_PROMPT, '');
        } else {
            return await stepContext.next(siteDetails.title);
        }
    }

    /**
     * If a description has not been provided, prompt for one.
     */
    private async descriptionStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        // Capture the results of the previous step
        siteDetails.title = stepContext.result;
        if (!siteDetails.description) {
            const promptText = `Provide a description for your ${siteDetails.siteType} site`;
            const descCard: Attachment = CardFactory.adaptiveCard(JSON.parse(
                JSON.stringify(GenericCard).replace('$Placeholder', promptText)));

            await stepContext.context.sendActivity({ attachments: [descCard] });
            return await stepContext.prompt('textPrompt', '');    
        } else {
            return await stepContext.next(siteDetails.description);
        }
    }

    /**
     * If an owner has not been provided, prompt for one.
     */
    private async ownerStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        // Capture the results of the previous step
        siteDetails.description = stepContext.result;

        if (!siteDetails.owner) {
            return await stepContext.beginDialog(OWNER_RESOLVER_DIALOG, { siteDetails });
        } else {
            return await stepContext.next(siteDetails.owner);
        }
    }

    /**
     * If an owner has not been provided, prompt for one.
     */
    private async aliasStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        // Capture the results of the previous step
        siteDetails.owner = stepContext.result;
        
        // Don't ask for alias if a communication site
        if (siteDetails.siteType === 'Communication') {
            
            return await stepContext.next();
        
        // Otherwise ask for an alias
        } else {
            
            if (!siteDetails.alias) {
                
                return await stepContext.beginDialog(ALIAS_RESOLVER_DIALOG, { siteDetails });   
            } else {
                return await stepContext.next(siteDetails.alias);
            }
        }
    }

    /**
     * Confirm the information the user has provided.
     */
    private async confirmStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        // Capture the results of the previous step
        siteDetails.alias = stepContext.result;
        
        const summaryCard: Attachment = CardFactory.adaptiveCard(JSON.parse(
            JSON.stringify(SummaryCard)
                .replace('$Title', siteDetails.title)
                .replace('$Desc', siteDetails.description)
                .replace('$Owner', siteDetails.owner)
                .replace('$Type', siteDetails.siteType)
                .replace('$Alias', siteDetails.alias ? siteDetails.alias : '' )
                ));

        await stepContext.context.sendActivity({ attachments: [summaryCard] });

        // Offer a YES/NO prompt.
        return await stepContext.prompt(CONFIRM_PROMPT, { prompt: '' });
    }

    /**
     * Complete the interaction and end the dialog.
     */
    private async finalStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        if (stepContext.result === true) {
            const siteDetails = stepContext.options as SiteDetails;

            return await stepContext.endDialog(siteDetails);
        } else {
            return await stepContext.endDialog();
        }
    }

}
