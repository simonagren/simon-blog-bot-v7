import { Attachment, CardFactory } from 'botbuilder';
import {
  ComponentDialog,
  DialogTurnResult,
  OAuthPrompt,
  PromptValidatorContext,
  TextPrompt,
  WaterfallDialog,
  WaterfallStepContext
} from 'botbuilder-dialogs';
import { GraphHelper } from '../helpers/graphHelper';
import GenericCard from '../resources/generic.json';

const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';
const OAUTH_PROMPT = 'OAuthPrompt';

export class AliasResolverDialog extends ComponentDialog {
  private static tokenResponse: any;
  
  private static async aliasPromptValidator(promptContext: PromptValidatorContext<string>): Promise<boolean> {
    if (promptContext.recognized.succeeded) {
      
      const alias: string = promptContext.recognized.value;
      
      if (await GraphHelper.aliasExistsPnP(alias, AliasResolverDialog.tokenResponse))  {
        promptContext.context.sendActivity('Alias already exist.');
        return false;
      }

      return true;

    } else {
      return false;
    }
  }

  constructor(id: string) {
    super(id || 'ownerResolverDialog');
    
    this
        .addDialog(new TextPrompt(TEXT_PROMPT, AliasResolverDialog.aliasPromptValidator.bind(this)))
        .addDialog(new OAuthPrompt(OAUTH_PROMPT, {
          connectionName: process.env.connectionName,
          text: 'Please Sign In',
          timeout: 300000,
          title: 'Sign In'
        }))
        .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
          this.promptStep.bind(this),
          this.initialStep.bind(this),
          this.finalStep.bind(this)
        ]));

    this.initialDialogId = WATERFALL_DIALOG;

  }

  /**
   * Prompt step in the waterfall. 
   */
  private async promptStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
      return await stepContext.beginDialog(OAUTH_PROMPT);
  }

  private async initialStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
    const tokenResponse = stepContext.result;
    if (tokenResponse && tokenResponse.token) {
      
      AliasResolverDialog.tokenResponse = tokenResponse;

      const siteDetails = (stepContext.options as any).siteDetails;
      const promptMsg = `Provide an alias for your ${siteDetails.siteType} site`;

      if (!siteDetails.alias) {

        const aliasCard: Attachment = CardFactory.adaptiveCard(JSON.parse(
          JSON.stringify(GenericCard).replace('$Placeholder', promptMsg)));
  
        await stepContext.context.sendActivity({ attachments: [aliasCard] });
        return await stepContext.prompt(TEXT_PROMPT,
          {
              prompt: ''
          });
      } else {
        return await stepContext.next(siteDetails.alias);
      }
    }
    await stepContext.context.sendActivity('Login was not successful please try again.');
    return await stepContext.endDialog();
  }

  private async finalStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {    
    const owner = stepContext.result;
    return await stepContext.endDialog(owner);
  }
}
