import React, { useEffect, useRef } from 'react';
import * as AdaptiveCards from 'adaptivecards';

export default function AdaptiveCardViewer({ cardJson, onExecuteAction, theme = 'teams' }) {
  const containerRef = useRef(null);

  useEffect(() => {
    if (!containerRef.current || !cardJson) return;

    // Clear previous rendered card
    containerRef.current.innerHTML = '';

    // Create an AdaptiveCard instance
    const adaptiveCard = new AdaptiveCards.AdaptiveCard();

    // Set Host Config based on Teams or Outlook look and feel
    adaptiveCard.hostConfig = new AdaptiveCards.HostConfig({
      fontFamily: 'Segoe UI, Helvetica Neue, sans-serif',
      spacing: {
        small: 4,
        default: 8,
        medium: 12,
        large: 16,
        extraLarge: 24,
        padding: 16
      },
      separator: {
        lineThickness: 1,
        lineColor: '#e0e0e0'
      },
      supportsInteractivity: true,
      containerStyles: {
        default: {
          backgroundColor: theme === 'teams' ? '#ffffff' : '#fafafa',
          foregroundColors: {
            default: { default: '#242424', subtle: '#616161' },
            accent: { default: '#5b5fc7', subtle: '#7b83eb' },
            good: { default: '#237b4b', subtle: '#34995f' },
            warning: { default: '#b26200', subtle: '#cc7700' },
            attention: { default: '#c4314b', subtle: '#e54d68' }
          }
        },
        emphasis: {
          backgroundColor: theme === 'teams' ? '#f5f5f7' : '#f0f4f8',
          foregroundColors: {
            default: { default: '#242424', subtle: '#616161' }
          }
        },
        good: {
          backgroundColor: '#eaf5ea',
          foregroundColors: { default: '#107c41' }
        },
        attention: {
          backgroundColor: '#fde7e9',
          foregroundColors: { default: '#a80000' }
        }
      },
      actions: {
        maxActions: 5,
        spacing: 'Default',
        buttonSpacing: 10,
        showCard: {
          actionMode: 'Inline'
        },
        actionsOrientation: 'Horizontal',
        actionAlignment: 'Left'
      }
    });

    // Handle button action events (Action.Submit, Action.OpenUrl)
    adaptiveCard.onExecuteAction = (action) => {
      if (action instanceof AdaptiveCards.SubmitAction) {
        // Collect all inputs and action data payload
        const inputsAndData = { ...action.data, ...adaptiveCard.getAllInputs() };
        if (onExecuteAction) {
          onExecuteAction(inputsAndData);
        }
      } else if (action instanceof AdaptiveCards.OpenUrlAction) {
        window.open(action.url, '_blank', 'noopener,noreferrer');
      }
    };

    try {
      // Parse JSON into card hierarchy
      adaptiveCard.parse(cardJson);
      // Render to standard HTML DOM elements
      const renderedCard = adaptiveCard.render();
      if (renderedCard) {
        containerRef.current.appendChild(renderedCard);
      }
    } catch (err) {
      console.error("Adaptive Card render error:", err);
      containerRef.current.innerHTML = `<div style="color:red; padding:10px;">Failed to render Adaptive Card: ${err.message}</div>`;
    }
  }, [cardJson, onExecuteAction, theme]);

  return (
    <div className={`card-preview-wrapper ${theme}-theme`}>
      <div ref={containerRef} className="adaptive-card-container" />
    </div>
  );
}

