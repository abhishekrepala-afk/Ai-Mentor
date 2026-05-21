import express from "express";
import Stripe from "stripe";
import {protect} from "../middleware/authMiddleware.js";

const router = express.Router();

// Initialize Stripe
const stripe = new Stripe(process.env.STRIPE_SECRET_KEY);

// CREATE CHECKOUT SESSION
router.post("/create-checkout-session", protect, async (req, res) => {
  try {
    const { course } = req.body;

    console.log("Incoming course:", course);

    // Validate course data
    if (
      !course ||
      !course.id ||
      !course.title ||
      !course.priceValue ||
      Number(course.priceValue) <= 0
    ) {
      return res.status(400).json({
        error: "Invalid course data",
      });
    }

    // Convert ₹ to paise
    const amount = Math.round(Number(course.priceValue) * 100);

    // Build success URL
    const successUrl = `${
      process.env.FRONTEND_URL
    }/success?courseId=${course.id}&title=${encodeURIComponent(
      course.title
    )}`;

    console.log("✅ SUCCESS URL:", successUrl);

    const session = await stripe.checkout.sessions.create({
      payment_method_types: ["card"],

      mode: "payment",

      line_items: [
        {
          price_data: {
            currency: "inr",
            product_data: {
              name: course.title,
            },
            unit_amount: amount,
          },
          quantity: 1,
        },
      ],

      // Important metadata
      metadata: {
        userId: req.user.id,
        courseId: course.id.toString(),
        courseTitle: course.title,
      },

      success_url: successUrl,

      cancel_url: `${process.env.FRONTEND_URL}/courses`,
    });

    return res.status(200).json({
      url: session.url,
    });
  } catch (error) {
    console.error("❌ Stripe Error:", error.message);

    return res.status(500).json({
      error: "Stripe session failed",
    });
  }
});

export default router;